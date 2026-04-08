# -*- coding: utf-8 -*-
"""Unit tests for Email MS Graph channel."""
from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from copaw.app.channels.email_ms_graph.channel import EmailMSGraphChannel
from copaw.app.channels.email_ms_graph.auth import MSGraphAuthManager
from copaw.app.channels.email_ms_graph.utils import (
    html_to_text,
    text_to_html,
    extract_sender_email,
    extract_conversation_id,
    should_process_message,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


@pytest.fixture
def mock_process():
    """Mock process handler."""

    async def _noop_process(_request):
        yield

    return _noop_process


@pytest.fixture
def mock_auth_manager():
    """Mock auth manager."""
    mock = MagicMock(spec=MSGraphAuthManager)
    mock.get_valid_token.return_value = "mock_token"
    mock.get_authorization_url.return_value = (
        "https://login.example.com/authorize"
    )
    return mock


@pytest.fixture
def sample_email_message():
    """Sample email message from MS Graph API."""
    return {
        "id": "AAMkAGI2THVSAAA=",
        "subject": "Test Email",
        "conversationId": "AAQkAGI2TH4=",
        "receivedDateTime": "2026-04-07T08:00:00Z",
        "from": {
            "emailAddress": {
                "name": "Test User",
                "address": "test@example.com",
            },
        },
        "body": {
            "contentType": "HTML",
            "content": "<html><body><p>Hello, this is a test.</p></body></html>",
        },
        "isRead": False,
        "hasAttachments": False,
    }


# ---------------------------------------------------------------------------
# Test Configuration Loading
# ---------------------------------------------------------------------------


class TestChannelConfiguration:
    """Test channel configuration and initialization."""

    @patch("copaw.app.channels.email_ms_graph.channel.MSGraphAuthManager")
    def test_from_config_dict(self, mock_auth_class, mock_process):
        """Test creating channel from dict config."""
        # Mock the auth manager constructor
        mock_auth_class.return_value = MagicMock(spec=MSGraphAuthManager)
        config = {
            "enabled": True,
            "tenant_id": "test-tenant",
            "client_id": "test-client",
            "client_secret": "test-secret",
            "redirect_uri": "http://localhost:8080/callback",
            "receive_mode": "polling",
            "poll_interval_sec": 30.0,
            "allowed_senders": ["user@example.com"],
            "subject_prefix": "[CoPaw]",
        }

        channel = EmailMSGraphChannel.from_config(
            process=mock_process,
            config=config,
        )

        assert channel.channel == "email_ms_graph"
        assert channel.receive_mode == "polling"
        assert channel.poll_interval_sec == 30.0
        assert "user@example.com" in channel.allowed_senders
        assert channel.subject_prefix == "[CoPaw]"

    @patch("copaw.app.channels.email_ms_graph.channel.MSGraphAuthManager")
    def test_channel_identifier(self, mock_auth_class, mock_process):
        """Test channel has correct identifier."""
        mock_auth_class.return_value = MagicMock(spec=MSGraphAuthManager)
        channel = EmailMSGraphChannel(
            auth_mode="delegated",
            mailbox_id="me",
            process=mock_process,
            tenant_id="test",
            client_id="test",
            client_secret="test",
            redirect_uri="http://test",
        )

        assert channel.channel == "email_ms_graph"
        assert channel.uses_manager_queue is True


# ---------------------------------------------------------------------------
# Test Utility Functions
# ---------------------------------------------------------------------------


class TestUtils:
    """Test utility functions."""

    def test_html_to_text(self):
        """Test HTML to text conversion."""
        html = "<html><body><p>Hello</p><br><p>World</p></body></html>"
        text = html_to_text(html)
        assert "Hello" in text
        assert "World" in text
        assert "<" not in text

    def test_html_to_text_with_entities(self):
        """Test HTML entity unescaping."""
        html = "<p>Tom &amp; Jerry &lt;test@example.com&gt;</p>"
        text = html_to_text(html)
        assert "Tom & Jerry" in text
        assert "<test@example.com>" in text

    def test_text_to_html(self):
        """Test text to HTML conversion."""
        text = "Hello\nWorld"
        html = text_to_html(text)
        assert "Hello<br>" in html
        assert "World" in html
        assert "<html>" in html

    def test_extract_sender_email(self, sample_email_message):
        """Test extracting sender email."""
        email = extract_sender_email(sample_email_message)
        assert email == "test@example.com"

    def test_extract_conversation_id(self, sample_email_message):
        """Test extracting conversation ID."""
        conv_id = extract_conversation_id(sample_email_message)
        assert conv_id == "AAQkAGI2TH4="

    def test_should_process_message_no_filters(self, sample_email_message):
        """Test message processing with no filters."""
        should_process = should_process_message(
            sample_email_message,
            allowed_senders=set(),
            subject_prefix="",
        )
        assert should_process is True

    def test_should_process_message_allowed_sender(self, sample_email_message):
        """Test message processing with allowed sender filter."""
        # Should process - sender in allowlist
        should_process = should_process_message(
            sample_email_message,
            allowed_senders={"test@example.com"},
            subject_prefix="",
        )
        assert should_process is True

        # Should not process - sender not in allowlist
        should_process = should_process_message(
            sample_email_message,
            allowed_senders={"other@example.com"},
            subject_prefix="",
        )
        assert should_process is False

    def test_should_process_message_subject_prefix(self, sample_email_message):
        """Test message processing with subject prefix filter."""
        # Should process - has prefix
        sample_email_message["subject"] = "[CoPaw] Test Email"
        should_process = should_process_message(
            sample_email_message,
            allowed_senders=set(),
            subject_prefix="[CoPaw]",
        )
        assert should_process is True

        # Should not process - no prefix
        sample_email_message["subject"] = "Test Email"
        should_process = should_process_message(
            sample_email_message,
            allowed_senders=set(),
            subject_prefix="[CoPaw]",
        )
        assert should_process is False


# ---------------------------------------------------------------------------
# Test Session ID Resolution
# ---------------------------------------------------------------------------


class TestSessionIdResolution:
    """Test session ID resolution."""

    @patch("copaw.app.channels.email_ms_graph.channel.MSGraphAuthManager")
    def test_resolve_session_id_with_conversation(
        self,
        mock_auth_class,
        mock_process,
    ):
        """Test session ID includes conversation ID."""
        mock_auth_class.return_value = MagicMock(spec=MSGraphAuthManager)
        channel = EmailMSGraphChannel(
            auth_mode="delegated",
            mailbox_id="me",
            process=mock_process,
            tenant_id="test",
            client_id="test",
            client_secret="test",
            redirect_uri="http://test",
        )

        session_id = channel.resolve_session_id(
            "user@example.com",
            {"conversation_id": "conv123"},
        )

        assert session_id == "email_ms_graph::user@example.com::conv123"

    @patch("copaw.app.channels.email_ms_graph.channel.MSGraphAuthManager")
    def test_resolve_session_id_without_conversation(
        self,
        mock_auth_class,
        mock_process,
    ):
        """Test session ID fallback without conversation ID."""
        mock_auth_class.return_value = MagicMock(spec=MSGraphAuthManager)
        channel = EmailMSGraphChannel(
            auth_mode="delegated",
            mailbox_id="me",
            process=mock_process,
            tenant_id="test",
            client_id="test",
            client_secret="test",
            redirect_uri="http://test",
        )

        session_id = channel.resolve_session_id(
            "user@example.com",
            {},
        )

        assert session_id == "email_ms_graph::user@example.com"


# ---------------------------------------------------------------------------
# Test Message Processing
# ---------------------------------------------------------------------------


class TestMessageProcessing:
    """Test email message processing."""

    @patch("copaw.app.channels.email_ms_graph.channel.MSGraphAuthManager")
    @pytest.mark.asyncio
    async def test_process_email_enqueues_payload(
        self,
        mock_auth_class,
        mock_process,
        sample_email_message,
    ):
        """Test processing email enqueues payload."""
        mock_auth_class.return_value = MagicMock(spec=MSGraphAuthManager)
        channel = EmailMSGraphChannel(
            auth_mode="delegated",
            mailbox_id="me",
            process=mock_process,
            tenant_id="test",
            client_id="test",
            client_secret="test",
            redirect_uri="http://test",
        )

        # Mock enqueue callback
        enqueue_mock = MagicMock()
        channel.set_enqueue(enqueue_mock)

        # Mock graph client mark_as_read
        channel.graph_client.mark_as_read = AsyncMock(return_value=True)

        # Process email
        await channel._process_email(sample_email_message)

        # Verify enqueue was called
        assert enqueue_mock.called
        payload = enqueue_mock.call_args[0][0]

        assert payload["channel_id"] == "email_ms_graph"
        assert payload["sender_id"] == "test@example.com"
        assert len(payload["content_parts"]) > 0
        assert payload["meta"]["message_id"] == "AAMkAGI2THVSAAA="
        assert payload["meta"]["subject"] == "Test Email"

    @patch("copaw.app.channels.email_ms_graph.channel.MSGraphAuthManager")
    @pytest.mark.asyncio
    async def test_process_email_converts_html_to_text(
        self,
        mock_auth_class,
        mock_process,
        sample_email_message,
    ):
        """Test HTML email body is converted to text."""
        mock_auth_class.return_value = MagicMock(spec=MSGraphAuthManager)
        channel = EmailMSGraphChannel(
            auth_mode="delegated",
            mailbox_id="me",
            process=mock_process,
            tenant_id="test",
            client_id="test",
            client_secret="test",
            redirect_uri="http://test",
        )

        enqueue_mock = MagicMock()
        channel.set_enqueue(enqueue_mock)
        channel.graph_client.mark_as_read = AsyncMock(return_value=True)

        await channel._process_email(sample_email_message)

        payload = enqueue_mock.call_args[0][0]
        content_parts = payload["content_parts"]

        # Check that HTML was converted
        text_part = content_parts[0]
        assert "Hello, this is a test" in text_part.text
        assert "<html>" not in text_part.text


# ---------------------------------------------------------------------------
# Test AgentRequest Building
# ---------------------------------------------------------------------------


class TestAgentRequestBuilding:
    """Test building AgentRequest from email payload."""

    @patch("copaw.app.channels.email_ms_graph.channel.MSGraphAuthManager")
    def test_build_agent_request_from_native(
        self,
        mock_auth_class,
        mock_process,
    ):
        """Test building AgentRequest from native email payload."""
        mock_auth_class.return_value = MagicMock(spec=MSGraphAuthManager)
        channel = EmailMSGraphChannel(
            auth_mode="delegated",
            mailbox_id="me",
            process=mock_process,
            tenant_id="test",
            client_id="test",
            client_secret="test",
            redirect_uri="http://test",
        )

        from agentscope_runtime.engine.schemas.agent_schemas import (
            TextContent,
            ContentType,
        )

        payload = {
            "channel_id": "email_ms_graph",
            "sender_id": "test@example.com",
            "content_parts": [
                TextContent(type=ContentType.TEXT, text="Hello"),
            ],
            "meta": {
                "conversation_id": "conv123",
                "message_id": "msg123",
            },
        }

        request = channel.build_agent_request_from_native(payload)

        assert request is not None
        assert request.channel == "email_ms_graph"
        assert request.user_id == "test@example.com"
        assert request.session_id and "conv123" in request.session_id


# ---------------------------------------------------------------------------
# Test Module Import
# ---------------------------------------------------------------------------


def test_module_import():
    """Test that the module can be imported without errors."""
    from copaw.app.channels.email_ms_graph import (
        EmailMSGraphChannel as ImportedChannel,
    )
    from copaw.app.channels.email_ms_graph.auth import MSGraphAuthManager
    from copaw.app.channels.email_ms_graph.graph_client import MSGraphClient

    assert ImportedChannel is not None
    assert MSGraphAuthManager is not None
    assert MSGraphClient is not None


def test_channel_registered():
    """Test that channel is registered in registry."""
    from copaw.app.channels.registry import get_channel_registry

    registry = get_channel_registry()
    assert "email_ms_graph" in registry
    assert registry["email_ms_graph"].__name__ == "EmailMSGraphChannel"
