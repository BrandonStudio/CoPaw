# -*- coding: utf-8 -*-
"""Email channel using Microsoft Graph API."""
from __future__ import annotations

import asyncio
import logging
import re
from collections import deque
from datetime import datetime, timedelta, timezone
from typing import Any, Optional, Union

from agentscope_runtime.engine.schemas.agent_schemas import (
    TextContent,
    ContentType,
)

from ....config.config import EmailMSGraphConfig
from ..base import (
    BaseChannel,
    OnReplySent,
    ProcessHandler,
)
from .auth import MSGraphAuthManager
from .graph_client import MSGraphClient
from .webhook import WebhookManager
from .utils import (
    html_to_text,
    text_to_html,
    extract_sender_email,
    extract_sender_name,
    extract_conversation_id,
    should_process_message,
)

logger = logging.getLogger(__name__)


class EmailMSGraphChannel(BaseChannel):
    """Email channel using Microsoft Graph API.

    Supports two receive modes:
    - Polling: Periodically check inbox for new messages
    - Webhook: Receive real-time notifications via MS Graph subscriptions
    """

    channel = "email_ms_graph"
    uses_manager_queue = True

    def __init__(
        self,
        process: ProcessHandler,
        auth_mode: str,
        mailbox_id: str,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        redirect_uri: str,
        receive_mode: str = "polling",
        poll_interval_sec: float = 60.0,
        webhook_url: str = "",
        webhook_notification_path: str = "",
        allowed_senders: Optional[list] = None,
        subject_prefix: str = "",
        on_reply_sent: OnReplySent = None,
        show_tool_details: bool = True,
        filter_tool_messages: bool = False,
        filter_thinking: bool = False,
        dm_policy: str = "open",
        group_policy: str = "open",
        allow_from: Optional[list] = None,
        deny_message: str = "",
    ):
        """Initialize email channel."""
        super().__init__(
            process,
            on_reply_sent=on_reply_sent,
            show_tool_details=show_tool_details,
            filter_tool_messages=filter_tool_messages,
            filter_thinking=filter_thinking,
            dm_policy=dm_policy,
            group_policy=group_policy,
            allow_from=allow_from,
            deny_message=deny_message,
        )

        self.auth_manager = MSGraphAuthManager(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
            redirect_uri=redirect_uri,
            auth_mode=auth_mode,
        )

        self.graph_client = MSGraphClient(
            auth_manager=self.auth_manager,
            mailbox_id=mailbox_id,
        )
        self.receive_mode = receive_mode
        self.poll_interval_sec = poll_interval_sec
        self.webhook_url = webhook_url
        self.allowed_senders = set(allowed_senders or [])
        self.subject_prefix = subject_prefix

        self._running = False
        self._poll_task: Optional[asyncio.Task] = None
        self._webhook_manager: Optional[WebhookManager] = None
        self._last_check_time: Optional[datetime] = None
        self._processed_message_ids: deque[str] = deque(maxlen=1000)

    @classmethod
    def from_config(
        cls,
        process: ProcessHandler,
        config: Union[EmailMSGraphConfig, dict],
        on_reply_sent: OnReplySent = None,
        show_tool_details: bool = True,
        filter_tool_messages: bool = False,
        filter_thinking: bool = False,
    ) -> "EmailMSGraphChannel":
        """Create channel from configuration."""
        if isinstance(config, dict):
            return cls(
                process=process,
                auth_mode=config.get("auth_mode", "delegated"),
                mailbox_id=config.get("mailbox_id", "me"),
                tenant_id=config.get("tenant_id", ""),
                client_id=config.get("client_id", ""),
                client_secret=config.get("client_secret", ""),
                redirect_uri=config.get(
                    "redirect_uri",
                    "http://localhost:8080/api/channels/email_ms_graph/callback",
                ),
                receive_mode=config.get("receive_mode", "polling"),
                poll_interval_sec=float(config.get("poll_interval_sec", 60.0)),
                webhook_url=config.get("webhook_url", ""),
                webhook_notification_path=config.get(
                    "webhook_notification_path",
                    "",
                ),
                allowed_senders=config.get("allowed_senders", []),
                subject_prefix=config.get("subject_prefix", ""),
                on_reply_sent=on_reply_sent,
                show_tool_details=show_tool_details,
                filter_tool_messages=filter_tool_messages,
                filter_thinking=filter_thinking,
                dm_policy=config.get("dm_policy", "open"),
                group_policy=config.get("group_policy", "open"),
                allow_from=config.get("allow_from"),
                deny_message=config.get("deny_message", ""),
            )
        else:
            return cls(
                process=process,
                auth_mode=config.auth_mode,
                mailbox_id=config.mailbox_id,
                tenant_id=config.tenant_id,
                client_id=config.client_id,
                client_secret=config.client_secret,
                redirect_uri=config.redirect_uri,
                receive_mode=config.receive_mode,
                poll_interval_sec=config.poll_interval_sec,
                webhook_url=config.webhook_url,
                webhook_notification_path=config.webhook_notification_path,
                allowed_senders=config.allowed_senders,
                subject_prefix=config.subject_prefix,
                on_reply_sent=on_reply_sent,
                show_tool_details=show_tool_details,
                filter_tool_messages=filter_tool_messages,
                filter_thinking=filter_thinking,
                dm_policy=config.dm_policy,
                group_policy=config.group_policy,
                allow_from=config.allow_from,
                deny_message=config.deny_message,
            )

    async def start(self) -> None:
        """Start the channel."""
        if self._running:
            logger.warning("Email channel already running")
            return

        logger.info(
            "Starting email MS Graph channel in %s mode",
            self.receive_mode,
        )

        # Get access token
        token = self.auth_manager.get_valid_token()
        if not token:
            if self.auth_manager.auth_mode == "delegated":
                logger.error(
                    "No valid access token. Please authorize at: %s",
                    self.auth_manager.get_authorization_url(),
                )
            else:
                logger.error(
                    "Failed to acquire app-only token. Check client credentials.",
                )
            return

        self._running = True

        if self.receive_mode == "webhook":
            await self._start_webhook_mode()
        else:
            await self._start_polling_mode()

    async def stop(self) -> None:
        """Stop the channel."""
        if not self._running:
            return

        logger.info("Stopping email MS Graph channel")
        self._running = False

        if self.receive_mode == "webhook" and self._webhook_manager:
            await self._webhook_manager.stop()
            self._webhook_manager = None
        elif self._poll_task:
            self._poll_task.cancel()
            try:
                await self._poll_task
            except asyncio.CancelledError:
                pass
            self._poll_task = None

        await self.graph_client.close()
        logger.info("Email MS Graph channel stopped")

    async def _start_polling_mode(self) -> None:
        """Start polling mode."""
        logger.info(
            "Starting polling mode with interval: %s seconds",
            self.poll_interval_sec,
        )
        self._last_check_time = datetime.now(timezone.utc) - timedelta(hours=1)
        self._poll_task = asyncio.create_task(self._poll_loop())

    async def _start_webhook_mode(self) -> None:
        """Start webhook mode."""
        if not self.webhook_url:
            logger.error(
                "webhook_url not configured, cannot start webhook mode",
            )
            self._running = False
            return

        logger.info("Starting webhook mode with URL: %s", self.webhook_url)

        self._webhook_manager = WebhookManager(
            graph_client=self.graph_client,
            notification_url=self.webhook_url,
            on_notification=self._handle_webhook_notification,
        )

        success = await self._webhook_manager.start()
        if not success:
            logger.error("Failed to start webhook manager")
            self._running = False

    async def _poll_loop(self) -> None:
        """Polling loop."""
        while self._running:
            try:
                await self._fetch_new_messages()
            except Exception as e:
                logger.exception("Error in polling loop: %s", e)

            await asyncio.sleep(self.poll_interval_sec)

    async def _fetch_new_messages(self) -> None:
        """Fetch new messages."""
        filter_query = None
        if self._last_check_time:
            time_str = self._last_check_time.isoformat()
            filter_query = f"receivedDateTime gt {time_str}"

        logger.debug("Fetching new messages with filter: %s", filter_query)

        messages = await self.graph_client.list_messages(
            folder="inbox",
            filter_query=filter_query,
            top=50,
        )

        logger.debug("Found %s messages", len(messages))

        for message in messages:
            message_id = message.get("id", "")
            if not message_id or message_id in self._processed_message_ids:
                continue

            if not should_process_message(
                message,
                self.allowed_senders,
                self.subject_prefix,
            ):
                logger.debug(
                    "Skipping message %s due to filters",
                    message_id[:20],
                )
                self._processed_message_ids.append(message_id)
                continue

            await self._process_email(message)
            self._processed_message_ids.append(message_id)

        self._last_check_time = datetime.now(timezone.utc)

        if len(self._processed_message_ids) > 1000:
            self._processed_message_ids = deque(
                list(self._processed_message_ids)[-1000:],
            )

    def _handle_webhook_notification(
        self,
        notification: dict[str, Any],
    ) -> None:
        """Handle webhook notification."""
        resource = notification.get("resource", "")
        logger.info("Received webhook notification for resource: %s", resource)

        match = re.search(r"messages\('([^']+)'\)", resource)
        if not match:
            logger.warning(
                "Could not extract message ID from resource: %s",
                resource,
            )
            return

        message_id = match.group(1)
        asyncio.create_task(self._fetch_and_process_message(message_id))

    async def _fetch_and_process_message(self, message_id: str) -> None:
        """Fetch and process message."""
        if message_id in self._processed_message_ids:
            logger.debug("Message %s already processed", message_id[:20])
            return

        message = await self.graph_client.get_message(message_id)
        if not message:
            logger.warning("Could not fetch message %s", message_id[:20])
            return

        if not should_process_message(
            message,
            self.allowed_senders,
            self.subject_prefix,
        ):
            logger.debug("Skipping message %s due to filters", message_id[:20])
            self._processed_message_ids.append(message_id)
            return

        await self._process_email(message)
        self._processed_message_ids.append(message_id)

    async def _process_email(self, message: dict[str, Any]) -> None:
        """Process an email message."""
        message_id = message.get("id", "")
        subject = message.get("subject", "") or "(No Subject)"
        sender_email = extract_sender_email(message)
        sender_name = extract_sender_name(message)

        logger.info(
            "Processing email: subject='%s', from=%s <%s>",
            subject[:50],
            sender_name,
            sender_email,
        )

        body = message.get("body", {})
        content_type = body.get("contentType", "text").lower()
        content_value = body.get("content", "") or ""

        if content_type == "html":
            text_content = html_to_text(content_value)
        else:
            text_content = content_value

        content_parts = [
            TextContent(type=ContentType.TEXT, text=text_content),
        ]

        # TODO: Extract and process attachments

        conversation_id = extract_conversation_id(message)

        payload = {
            "channel_id": self.channel,
            "sender_id": sender_email,
            "content_parts": content_parts,
            "meta": {
                "message_id": message_id,
                "conversation_id": conversation_id,
                "subject": subject,
                "sender_name": sender_name,
                "sender_email": sender_email,
            },
        }

        if self._enqueue:
            self._enqueue(payload)
        else:
            logger.warning("No enqueue callback set")

        await self.graph_client.mark_as_read(message_id)

    def resolve_session_id(
        self,
        sender_id: str,
        channel_meta: dict[str, Any] | None = None,
    ) -> str:
        """Resolve session ID."""
        conversation_id = (
            channel_meta.get("conversation_id", "") if channel_meta else ""
        )
        if conversation_id:
            return f"{self.channel}::{sender_id}::{conversation_id}"
        else:
            return f"{self.channel}::{sender_id}"

    def build_agent_request_from_native(self, native_payload: Any):
        """Convert email payload to AgentRequest."""
        if not isinstance(native_payload, dict):
            logger.warning(
                "Invalid native payload type: %s",
                type(native_payload),
            )
            raise ValueError("Invalid native payload type")

        sender_id = native_payload.get("sender_id", "")
        content_parts = native_payload.get("content_parts", [])
        meta = native_payload.get("meta", {})

        session_id = self.resolve_session_id(sender_id, meta)

        request = self.build_agent_request_from_user_content(
            channel_id=self.channel,
            sender_id=sender_id,
            session_id=session_id,
            content_parts=content_parts,
            channel_meta=meta,
        )
        request.channel_meta = meta

        return request

    async def send_response(  # pylint: disable=too-many-branches
        self,
        to_handle: str,
        response: Any,
        meta: Optional[dict] = None,
    ) -> None:
        """Send agent response as email."""
        if not response:
            return

        meta = meta or {}
        original_message_id = meta.get("message_id")
        subject = meta.get("subject", "Re: Message from CoPaw")
        sender_email = meta.get("sender_email", "")

        output_messages = getattr(response, "output", [])
        if not output_messages:
            logger.debug("No output messages in response")
            return

        text_parts = []
        for msg in output_messages:
            content = getattr(msg, "content", [])
            for part in content:
                if isinstance(part, TextContent):
                    text = getattr(part, "text", "")
                    if text:
                        text_parts.append(text)

        if not text_parts:
            logger.debug("No text content to send")
            return

        body_text = "\n\n".join(text_parts)
        body_html = text_to_html(body_text)

        if original_message_id:
            logger.info(
                "Sending reply to message %s",
                original_message_id[:20],
            )
            success = await self.graph_client.send_reply(
                message_id=original_message_id,
                body_content=body_html,
            )
        else:
            logger.info("Sending new email to %s", sender_email)
            if not sender_email:
                logger.warning("No sender email, cannot send reply")
                return

            success = await self.graph_client.send_mail(
                to_recipients=[sender_email],
                subject=subject,
                body_content=body_html,
                body_type="HTML",
            )

        if success:
            logger.info("Email sent successfully")
            if self._on_reply_sent:
                session_id = getattr(response, "session_id", "")
                self._on_reply_sent(self.channel, sender_email, session_id)
        else:
            logger.error("Failed to send email")

    async def consume_one(self, payload: Any) -> None:
        """Consume one message from queue."""
        if isinstance(payload, dict):
            request = self.build_agent_request_from_native(payload)
        else:
            request = payload

        if not request:
            logger.warning("Could not build AgentRequest from payload")
            return

        await self._consume_with_tracker(request, payload)
