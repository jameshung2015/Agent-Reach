# -*- coding: utf-8 -*-
"""
SharePoint / OneDrive channel — support directory inspection and downloading.
"""

from .base import Channel


class SharePointChannel(Channel):
    name: str = "sharepoint"
    description: str = "SharePoint 与 OneDrive 文件/目录下载"
    backends = ["playwright", "requests"]
    tier = 0

    def can_handle(self, url: str) -> bool:
        lower = url.lower()
        return (
            "sharepoint.com" in lower
            or "sharepoint.cn" in lower
            or "1drv.ms" in lower
            or "onedrive.live.com" in lower
        )

    def check(self, config=None):
        try:
            import playwright
            import requests
            return "ok", "Playwright + Requests 可用，支持受限与密码保护链接下载"
        except ImportError as e:
            return "warn", f"缺少依赖: {e}. 请运行 pip install playwright requests"
