# -*- coding: utf-8 -*-
"""OAuth2 authentication manager for Microsoft Graph API."""
from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Optional, Dict, Any

import msal

logger = logging.getLogger(__name__)

# MS Graph API scopes needed for email operations
GRAPH_SCOPES = [
    "Mail.ReadWrite",
    "Mail.Send",
    "offline_access",  # For refresh token
]


class MSGraphAuthManager:
    """Manages OAuth2 authentication for Microsoft Graph API."""

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        redirect_uri: str,
        token_file: Optional[Path] = None,
    ):
        """Initialize auth manager.
        
        Args:
            tenant_id: Azure AD tenant ID or 'common' for multi-tenant
            client_id: Application (client) ID
            client_secret: Client secret value
            redirect_uri: OAuth2 redirect URI
            token_file: Path to persist token (default: working_dir/email_ms_graph_token.json)
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.redirect_uri = redirect_uri
        
        # Token storage
        if token_file is None:
            from copaw.constant import WORKING_DIR
            token_file = WORKING_DIR / "email_ms_graph_token.json"
        self.token_file = Path(token_file)
        
        # Build authority URL
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        
        # Create MSAL confidential client
        self.app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=authority,
        )
        
        self._cached_token: Optional[Dict[str, Any]] = None

    def get_authorization_url(self) -> str:
        """Generate OAuth2 authorization URL for user consent.
        
        Returns:
            Authorization URL to redirect user to
        """
        auth_url = self.app.get_authorization_request_url(
            scopes=GRAPH_SCOPES,
            redirect_uri=self.redirect_uri,
        )
        logger.info("Generated authorization URL: %s", auth_url)
        return auth_url

    def get_token_from_code(self, auth_code: str) -> Optional[Dict[str, Any]]:
        """Exchange authorization code for access token.
        
        Args:
            auth_code: Authorization code from OAuth2 callback
            
        Returns:
            Token response dict with access_token, refresh_token, etc.
            None if failed
        """
        try:
            result = self.app.acquire_token_by_authorization_code(
                code=auth_code,
                scopes=GRAPH_SCOPES,
                redirect_uri=self.redirect_uri,
            )
            
            if "access_token" in result:
                logger.info("Successfully acquired token from authorization code")
                self._cached_token = result
                self.save_token(result)
                return result
            else:
                error = result.get("error", "unknown_error")
                error_desc = result.get("error_description", "")
                logger.error(
                    "Failed to acquire token: %s - %s",
                    error,
                    error_desc,
                )
                return None
        except Exception as e:
            logger.exception("Error acquiring token from code: %s", e)
            return None

    def refresh_token(self) -> Optional[Dict[str, Any]]:
        """Refresh access token using refresh token.
        
        Returns:
            New token response or None if failed
        """
        # Load existing token
        token = self._cached_token or self.load_token()
        if not token:
            logger.warning("No token to refresh")
            return None
        
        refresh_token_str = token.get("refresh_token")
        if not refresh_token_str:
            logger.warning("No refresh token available")
            return None
        
        try:
            result = self.app.acquire_token_by_refresh_token(
                refresh_token=refresh_token_str,
                scopes=GRAPH_SCOPES,
            )
            
            if "access_token" in result:
                logger.info("Successfully refreshed access token")
                self._cached_token = result
                self.save_token(result)
                return result
            else:
                error = result.get("error", "unknown_error")
                error_desc = result.get("error_description", "")
                logger.error(
                    "Failed to refresh token: %s - %s",
                    error,
                    error_desc,
                )
                return None
        except Exception as e:
            logger.exception("Error refreshing token: %s", e)
            return None

    def get_valid_token(self) -> Optional[str]:
        """Get a valid access token, refreshing if necessary.
        
        Returns:
            Valid access token string or None if unable to get token
        """
        # Check if we have a cached token
        token = self._cached_token or self.load_token()
        if not token:
            logger.warning("No token available. User needs to authorize.")
            return None
        
        # Check if token is still valid (MSAL handles expiry internally)
        access_token = token.get("access_token")
        if not access_token:
            # Try to refresh
            logger.info("No access token, attempting refresh")
            token = self.refresh_token()
            if token:
                return token.get("access_token")
            return None
        
        # TODO: Check token expiry and refresh proactively
        # For now, rely on 401 responses to trigger refresh
        return access_token

    def save_token(self, token: Dict[str, Any]) -> None:
        """Persist token to disk.
        
        Args:
            token: Token response to save
        """
        try:
            self.token_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.token_file, "w", encoding="utf-8") as f:
                json.dump(token, f, indent=2)
            logger.info("Token saved to %s", self.token_file)
        except Exception as e:
            logger.exception("Failed to save token: %s", e)

    def load_token(self) -> Optional[Dict[str, Any]]:
        """Load token from disk.
        
        Returns:
            Token dict or None if not found
        """
        if not self.token_file.exists():
            logger.debug("Token file not found: %s", self.token_file)
            return None
        
        try:
            with open(self.token_file, "r", encoding="utf-8") as f:
                token = json.load(f)
            logger.info("Token loaded from %s", self.token_file)
            self._cached_token = token
            return token
        except Exception as e:
            logger.exception("Failed to load token: %s", e)
            return None

    def clear_token(self) -> None:
        """Clear cached token and delete token file."""
        self._cached_token = None
        if self.token_file.exists():
            try:
                self.token_file.unlink()
                logger.info("Token file deleted: %s", self.token_file)
            except Exception as e:
                logger.exception("Failed to delete token file: %s", e)
