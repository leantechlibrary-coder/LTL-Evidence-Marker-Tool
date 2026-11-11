"""テスト共通設定。PyQt は使わず stamp_core だけを叩く。

リポジトリ直下の stamp_core.py を import できるよう、親をパスに載せる。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
