"""Server-side screenshot generation for org chart using Playwright."""

from __future__ import annotations

import logging
import time
from typing import Optional

logger = logging.getLogger(__name__)

# Playwright is an optional dependency for screenshot generation
try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False
    logger.warning(
        "Playwright not available. Install with: pip install playwright && playwright install chromium"
    )


def is_playwright_available() -> bool:
    """Check if Playwright is installed and available."""
    return PLAYWRIGHT_AVAILABLE


def generate_org_chart_png_via_export(
    base_url: str,
    timeout_ms: int = 60000,
    auth_token: Optional[str] = None
) -> Optional[bytes]:
    """
    Generate PNG by triggering the client-side export function.
    
    This method uses the existing client-side PNG export functionality
    by programmatically triggering it and capturing the downloaded file.
    
    Args:
        base_url: The base URL of the application
        timeout_ms: Maximum time to wait for export
        auth_token: Optional authentication token
        
    Returns:
        PNG image as bytes, or None if generation failed
    """
    if not PLAYWRIGHT_AVAILABLE:
        logger.error("Playwright is not installed. Cannot generate PNG.")
        return None
    
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            
            context = browser.new_context(
                viewport={'width': 1920, 'height': 1080},
                accept_downloads=True
            )
            
            if auth_token:
                context.set_extra_http_headers({
                    'Authorization': f'Bearer {auth_token}'
                })
            
            page = context.new_page()
            
            # Navigate to the org chart page
            chart_url = f"{base_url.rstrip('/')}/"
            logger.info(f"Navigating to {chart_url} for PNG export")
            page.goto(chart_url, wait_until='networkidle', timeout=timeout_ms)
            
            # Wait for the chart to render
            page.wait_for_selector('svg', timeout=timeout_ms)
            time.sleep(2)
            
            # Trigger PNG export via JavaScript
            # This calls the existing exportToImage function
            with page.expect_download(timeout=timeout_ms) as download_info:
                page.evaluate("""
                    async () => {
                        // Call the existing export function with full chart option
                        await exportToImage('png', true);
                    }
                """)
            
            # Wait for download to complete and get the file
            download = download_info.value
            png_bytes = download.path()
            
            # Read the downloaded file
            with open(png_bytes, 'rb') as f:
                result = f.read()
            
            logger.info(f"PNG export generated successfully, size: {len(result)} bytes")
            
            browser.close()
            return result
            
    except PlaywrightTimeoutError as e:
        logger.error(f"Timeout while generating PNG export: {e}")
        return None
    except Exception as e:
        logger.error(f"Error generating PNG export: {e}", exc_info=True)
        return None
