"""
LibreOffice 子进程环境变量（最小配置）。

说明：部分沙箱环境会限制 AF_UNIX，需自行配置运行环境；本仓库不嵌入第三方 LD_PRELOAD shim。
"""

from __future__ import annotations

import os
from typing import Mapping


def get_soffice_env() -> Mapping[str, str]:
    """供 subprocess 调用 soffice 时合并到 env。"""
    env = dict(os.environ)
    # 无界面、无真实显示时常用，减少对 GUI 插件的依赖
    env.setdefault("SAL_USE_VCLPLUGIN", "svp")
    return env
