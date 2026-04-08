# -*- coding: utf-8 -*-
"""Webhook subscription management for MS Graph change notifications."""
from __future__ import annotations

import asyncio
import logging
from datetime import datetime, timezone
from typing import TYPE_CHECKING, Any, Callable

from .graph_client import MSGraphClient

if TYPE_CHECKING:
    from typing import Optional

logger = logging.getLogger(__name__)

# Subscription renewal interval (renew 1 hour before expiry)
RENEW_BEFORE_EXPIRY_MINUTES = 60


class WebhookManager:
    """Manages MS Graph webhook subscriptions for email notifications."""

    def __init__(
        self,
        graph_client: MSGraphClient,
        notification_url: str,
        on_notification: Callable[[dict[str, Any]], None],
    ):
        """Initialize webhook manager.

        Args:
            graph_client: MS Graph API client
            notification_url: Public URL for webhook notifications
            on_notification: Callback function when notification received
        """
        self.graph_client = graph_client
        self.notification_url = notification_url
        self.on_notification = on_notification

        self.subscription_id: Optional[str] = None
        self.expiration_time: Optional[datetime] = None
        self._renew_task: Optional[asyncio.Task] = None
        self._running = False

    async def create_subscription(self) -> bool:
        """Create a new webhook subscription.

        Returns:
            True if subscription created successfully
        """
        logger.info(
            "Creating webhook subscription for %s",
            self.notification_url,
        )

        result = await self.graph_client.create_subscription(
            notification_url=self.notification_url,
        )

        if result and "id" in result:
            self.subscription_id = result["id"]
            expiration_str = result.get("expirationDateTime", "")
            if expiration_str:
                try:
                    self.expiration_time = datetime.fromisoformat(
                        expiration_str.replace("Z", "+00:00"),
                    )
                except ValueError:
                    logger.warning(
                        "Failed to parse expiration time: %s",
                        expiration_str,
                    )

            logger.info(
                "Webhook subscription created: ID=%s, expires=%s",
                self.subscription_id,
                self.expiration_time,
            )
            return True
        else:
            logger.error("Failed to create webhook subscription")
            return False

    async def renew_subscription(self) -> bool:
        """Renew the current subscription.

        Returns:
            True if renewed successfully
        """
        if not self.subscription_id:
            logger.warning("No subscription to renew")
            return False

        logger.info("Renewing webhook subscription: %s", self.subscription_id)

        result = await self.graph_client.renew_subscription(
            subscription_id=self.subscription_id,
        )

        if result and "expirationDateTime" in result:
            expiration_str = result["expirationDateTime"]
            try:
                self.expiration_time = datetime.fromisoformat(
                    expiration_str.replace("Z", "+00:00"),
                )
            except ValueError:
                logger.warning(
                    "Failed to parse expiration time: %s",
                    expiration_str,
                )

            logger.info(
                "Webhook subscription renewed, expires=%s",
                self.expiration_time,
            )
            return True
        else:
            logger.error("Failed to renew webhook subscription")
            return False

    async def delete_subscription(self) -> bool:
        """Delete the current subscription.

        Returns:
            True if deleted successfully
        """
        if not self.subscription_id:
            logger.debug("No subscription to delete")
            return True

        logger.info("Deleting webhook subscription: %s", self.subscription_id)

        result = await self.graph_client.delete_subscription(
            subscription_id=self.subscription_id,
        )

        if result:
            logger.info("Webhook subscription deleted")
            self.subscription_id = None
            self.expiration_time = None
            return True
        else:
            logger.error("Failed to delete webhook subscription")
            return False

    async def _auto_renew_loop(self) -> None:
        """Automatic renewal loop - renews subscription before it expires."""
        while self._running:
            if not self.expiration_time:
                # No expiration time, wait and check again
                await asyncio.sleep(60)
                continue

            now = datetime.now(timezone.utc)
            time_until_expiry = (self.expiration_time - now).total_seconds()
            renew_in_seconds = time_until_expiry - (
                RENEW_BEFORE_EXPIRY_MINUTES * 60
            )

            if renew_in_seconds <= 0:
                # Time to renew
                logger.info("Subscription expiring soon, renewing...")
                success = await self.renew_subscription()
                if not success:
                    logger.error("Failed to renew subscription, will retry")
                    await asyncio.sleep(60)  # Retry in 1 minute
                    continue
            else:
                # Wait until renewal time
                logger.debug(
                    "Subscription renewal scheduled in %s seconds",
                    renew_in_seconds,
                )
                await asyncio.sleep(
                    min(renew_in_seconds, 3600),
                )  # Check at least hourly

    async def start(self) -> bool:
        """Start webhook manager (create subscription and auto-renewal).

        Returns:
            True if started successfully
        """
        if self._running:
            logger.warning("Webhook manager already running")
            return True

        # Create subscription
        success = await self.create_subscription()
        if not success:
            return False

        # Start auto-renewal loop
        self._running = True
        self._renew_task = asyncio.create_task(self._auto_renew_loop())
        logger.info("Webhook manager started")
        return True

    async def stop(self) -> None:
        """Stop webhook manager (cancel auto-renewal and delete subscription)."""
        if not self._running:
            return

        # Stop auto-renewal
        self._running = False
        if self._renew_task:
            self._renew_task.cancel()
            try:
                await self._renew_task
            except asyncio.CancelledError:
                pass
            self._renew_task = None

        # Delete subscription
        await self.delete_subscription()
        logger.info("Webhook manager stopped")

    def validate_notification(
        self,
        notification_data: dict[str, Any],
        expected_client_state: str = "CoPawEmailChannel",
    ) -> bool:
        """Validate webhook notification.

        Args:
            notification_data: Notification payload from MS Graph
            expected_client_state: Expected clientState value

        Returns:
            True if notification is valid
        """
        # Check clientState
        client_state = notification_data.get("clientState", "")
        if client_state != expected_client_state:
            logger.warning(
                "Invalid clientState in notification: %s",
                client_state,
            )
            return False

        # Additional validation can be added here
        # MS Graph doesn't use HMAC signatures like some other services,
        # but validates using clientState and HTTPS

        return True

    def handle_notification(self, notification_data: dict[str, Any]) -> None:
        """Handle incoming webhook notification.

        Args:
            notification_data: Notification payload
        """
        if not self.validate_notification(notification_data):
            logger.warning("Received invalid notification, ignoring")
            return

        logger.info("Received valid webhook notification")

        # Extract notification details
        value_list = notification_data.get("value", [])
        for item in value_list:
            change_type = item.get("changeType", "")
            resource = item.get("resource", "")
            _ = item.get("resourceData", {})  # Reserved for future use

            logger.debug(
                "Notification: changeType=%s, resource=%s",
                change_type,
                resource,
            )

            # Call the callback
            if self.on_notification:
                try:
                    self.on_notification(item)
                except Exception as e:
                    logger.exception("Error in notification callback: %s", e)
