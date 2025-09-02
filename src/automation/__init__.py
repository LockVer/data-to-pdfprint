"""
自动化模块

提供自动化工作流程功能
"""

from .workflow import AutomatedWorkflow, create_default_config
from ..data.data_processor import PackagingConfig

__all__ = ['AutomatedWorkflow', 'PackagingConfig', 'create_default_config']