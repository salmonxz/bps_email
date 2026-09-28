"""
Core package for Automated Provincial Email Dispatcher (BPS).
"""

from core.history_manager import HistoryManager, get_default_db_path

__all__ = ["HistoryManager", "get_default_db_path"]
