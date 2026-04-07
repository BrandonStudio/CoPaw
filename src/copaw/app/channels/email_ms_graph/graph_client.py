# -*- coding: utf-8 -*-
"""Microsoft Graph API client for email operations."""
from __future__ import annotations

import asyncio
import logging
from typing import Optional, Dict, Any, List
from datetime import datetime, timezone

import httpx

from .auth import MSGraphAuthManager

logger = logging.getLogger(__name__)

GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"
DEFAULT_TIMEOUT = 30.0


class MSGraphClient:
    """Client for Microsoft Graph API email operations."""

    def __init__(
        self,
        auth_manager: MSGraphAuthManager,
        timeout: float = DEFAULT_TIMEOUT,
    ):
        """Initialize Graph API client.
        
        Args:
            auth_manager: Authentication manager instance
            timeout: HTTP request timeout in seconds
        """
        self.auth = auth_manager
        self.timeout = timeout
        self._http_client: Optional[httpx.AsyncClient] = None

    async def _get_http_client(self) -> httpx.AsyncClient:
        """Get or create async HTTP client."""
        if self._http_client is None or self._http_client.is_closed:
            self._http_client = httpx.AsyncClient(timeout=self.timeout)
        return self._http_client

    async def close(self) -> None:
        """Close HTTP client."""
        if self._http_client and not self._http_client.is_closed:
            await self._http_client.aclose()

    async def _make_request(
        self,
        method: str,
        endpoint: str,
        json_data: Optional[Dict[str, Any]] = None,
        params: Optional[Dict[str, Any]] = None,
        retry_count: int = 1,
    ) -> Optional[Dict[str, Any]]:
        """Make HTTP request to Graph API with retry logic.
        
        Args:
            method: HTTP method (GET, POST, etc.)
            endpoint: API endpoint path (will be appended to base URL)
            json_data: JSON body for POST/PATCH requests
            params: Query parameters
            retry_count: Number of retries on failure
            
        Returns:
            Response JSON or None if failed
        """
        token = self.auth.get_valid_token()
        if not token:
            logger.error("No valid access token available")
            return None
        
        url = f"{GRAPH_API_ENDPOINT}{endpoint}"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }
        
        client = await self._get_http_client()
        
        for attempt in range(retry_count + 1):
            try:
                response = await client.request(
                    method=method,
                    url=url,
                    headers=headers,
                    json=json_data,
                    params=params,
                )
                
                if response.status_code == 401:
                    # Token expired, try to refresh
                    logger.info("Token expired (401), attempting refresh")
                    new_token_data = self.auth.refresh_token()
                    if new_token_data:
                        # Retry with new token
                        token = new_token_data.get("access_token")
                        headers["Authorization"] = f"Bearer {token}"
                        continue
                    else:
                        logger.error("Failed to refresh token")
                        return None
                
                if response.status_code == 429:
                    # Rate limited
                    retry_after = int(response.headers.get("Retry-After", "60"))
                    logger.warning("Rate limited, waiting %s seconds", retry_after)
                    if attempt < retry_count:
                        await asyncio.sleep(retry_after)
                        continue
                    return None
                
                if response.status_code >= 400:
                    logger.error(
                        "Graph API request failed: %s %s - Status: %s, Response: %s",
                        method,
                        endpoint,
                        response.status_code,
                        response.text[:500],
                    )
                    return None
                
                return response.json()
                
            except httpx.TimeoutException:
                logger.warning("Request timeout (attempt %s/%s)", attempt + 1, retry_count + 1)
                if attempt < retry_count:
                    await asyncio.sleep(2 ** attempt)  # Exponential backoff
                    continue
                return None
            except Exception as e:
                logger.exception("Error making Graph API request: %s", e)
                if attempt < retry_count:
                    await asyncio.sleep(2 ** attempt)
                    continue
                return None
        
        return None

    async def list_messages(
        self,
        folder: str = "inbox",
        filter_query: Optional[str] = None,
        select_fields: Optional[List[str]] = None,
        top: int = 10,
    ) -> List[Dict[str, Any]]:
        """List messages from a mail folder.
        
        Args:
            folder: Mail folder name (inbox, sentitems, etc.)
            filter_query: OData filter expression
            select_fields: List of fields to return
            top: Maximum number of messages to return
            
        Returns:
            List of message objects
        """
        endpoint = f"/me/mailFolders/{folder}/messages"
        params: Dict[str, Any] = {"$top": top, "$orderby": "receivedDateTime desc"}
        
        if filter_query:
            params["$filter"] = filter_query
        
        if select_fields:
            params["$select"] = ",".join(select_fields)
        
        result = await self._make_request("GET", endpoint, params=params)
        if result and "value" in result:
            return result["value"]
        return []

    async def get_message(self, message_id: str) -> Optional[Dict[str, Any]]:
        """Get a specific message by ID.
        
        Args:
            message_id: Message ID
            
        Returns:
            Message object or None if not found
        """
        endpoint = f"/me/messages/{message_id}"
        return await self._make_request("GET", endpoint)

    async def send_mail(
        self,
        to_recipients: List[str],
        subject: str,
        body_content: str,
        body_type: str = "HTML",
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
    ) -> bool:
        """Send a new email.
        
        Args:
            to_recipients: List of recipient email addresses
            subject: Email subject
            body_content: Email body content
            body_type: Body content type ("Text" or "HTML")
            cc_recipients: CC recipients
            bcc_recipients: BCC recipients
            
        Returns:
            True if sent successfully
        """
        message_data = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": body_type,
                    "content": body_content,
                },
                "toRecipients": [
                    {"emailAddress": {"address": addr}} for addr in to_recipients
                ],
            }
        }
        
        if cc_recipients:
            message_data["message"]["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in cc_recipients
            ]
        
        if bcc_recipients:
            message_data["message"]["bccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in bcc_recipients
            ]
        
        endpoint = "/me/sendMail"
        result = await self._make_request("POST", endpoint, json_data=message_data)
        return result is not None

    async def send_reply(
        self,
        message_id: str,
        body_content: str,
        body_type: str = "HTML",
    ) -> bool:
        """Reply to an existing message.
        
        Args:
            message_id: ID of message to reply to
            body_content: Reply body content
            body_type: Body content type ("Text" or "HTML")
            
        Returns:
            True if sent successfully
        """
        reply_data = {
            "comment": body_content,
            "message": {
                "body": {
                    "contentType": body_type,
                    "content": body_content,
                }
            }
        }
        
        endpoint = f"/me/messages/{message_id}/reply"
        result = await self._make_request("POST", endpoint, json_data=reply_data)
        return result is not None

    async def mark_as_read(self, message_id: str) -> bool:
        """Mark a message as read.
        
        Args:
            message_id: Message ID
            
        Returns:
            True if marked successfully
        """
        endpoint = f"/me/messages/{message_id}"
        data = {"isRead": True}
        result = await self._make_request("PATCH", endpoint, json_data=data)
        return result is not None

    async def create_subscription(
        self,
        notification_url: str,
        expiration_minutes: int = 4230,  # Max: 3 days (4320 min), use 4230 for safety
    ) -> Optional[Dict[str, Any]]:
        """Create a webhook subscription for email notifications.
        
        Args:
            notification_url: Public URL to receive webhook notifications
            expiration_minutes: Subscription duration in minutes (max 4320 = 3 days)
            
        Returns:
            Subscription object with id and expirationDateTime, or None if failed
        """
        # Calculate expiration time
        expiration_dt = datetime.now(timezone.utc)
        from datetime import timedelta
        expiration_dt += timedelta(minutes=expiration_minutes)
        
        subscription_data = {
            "changeType": "created",
            "notificationUrl": notification_url,
            "resource": "me/mailFolders('Inbox')/messages",
            "expirationDateTime": expiration_dt.isoformat(),
            "clientState": "CoPawEmailChannel",  # For validation
        }
        
        endpoint = "/subscriptions"
        return await self._make_request("POST", endpoint, json_data=subscription_data)

    async def renew_subscription(
        self,
        subscription_id: str,
        expiration_minutes: int = 4230,
    ) -> Optional[Dict[str, Any]]:
        """Renew an existing subscription.
        
        Args:
            subscription_id: Subscription ID to renew
            expiration_minutes: New expiration duration in minutes
            
        Returns:
            Updated subscription object or None if failed
        """
        expiration_dt = datetime.now(timezone.utc)
        from datetime import timedelta
        expiration_dt += timedelta(minutes=expiration_minutes)
        
        update_data = {
            "expirationDateTime": expiration_dt.isoformat(),
        }
        
        endpoint = f"/subscriptions/{subscription_id}"
        return await self._make_request("PATCH", endpoint, json_data=update_data)

    async def delete_subscription(self, subscription_id: str) -> bool:
        """Delete a subscription.
        
        Args:
            subscription_id: Subscription ID to delete
            
        Returns:
            True if deleted successfully
        """
        endpoint = f"/subscriptions/{subscription_id}"
        result = await self._make_request("DELETE", endpoint)
        return result is not None
