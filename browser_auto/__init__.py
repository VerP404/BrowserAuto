"""Общая автоматизация браузера через Firefox (geckodriver)."""

from browser_auto.auth import login_iszl, login_oms
from browser_auto.driver import GeckoBrowser

__all__ = ["GeckoBrowser", "login_iszl", "login_oms"]
