# -*- coding: utf-8 -*-
"""
Auto GUI Tool - Professional Automation Tool
Version 2.0

A professional-grade GUI automation tool with enhanced features,
improved error handling, and better user experience.
"""

# Standard library imports
import os
import sys
import time
import json
import logging
import threading
import subprocess
import uuid
import ctypes
from pathlib import Path
from datetime import datetime, timedelta
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional, Tuple, Any, Union
from contextlib import contextmanager
from collections import deque

# GUI library imports
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Third-party library imports
import cv2
import numpy as np
import pyautogui
import pyperclip
import mss
from screeninfo import get_monitors
from PIL import Image, ImageTk, ImageDraw

# ログ設定
class LogManager:
    @staticmethod
    def setup_logging():
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        
        log_file = log_dir / f"auto_gui_tool_{datetime.now().strftime('%Y%m%d')}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - [%(funcName)s:%(lineno)d] - %(message)s",
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        return logging.getLogger(__name__)

logger = LogManager.setup_logging()

# アプリケーション設定
class AppConfig:
    APP_NAME = "Auto GUI Tool Professional"
    VERSION = "2.0.0"
    DEFAULT_CONFIG_FILE = "config/last_config.json"
    DEFAULT_DPI_SCALE = 100.0
    
    # 🎨 3つの選択可能テーマ
    THEMES = {
        'dark_pro': {  # プロフェッショナル（現在のテーマ）
            'name': '🌃 Dark Professional',
            'bg_primary': '#1a1a1a',
            'bg_secondary': '#2d2d2d',
            'bg_tertiary': '#3a3a3a',
            'bg_quaternary': '#4a4a4a',
            'bg_accent': '#0078d4',
            'bg_hover': '#106ebe',
            'bg_pressed': '#005a9e',
            'fg_primary': '#ffffff',
            'fg_secondary': '#e1e1e1',
            'fg_tertiary': '#b3b3b3',
            'fg_accent': '#58c4dc',
            'success': '#16c60c',
            'warning': '#ff8c00',
            'error': '#e74c3c',
            'info': '#17a2b8',
            'border': '#404040',
            'border_light': '#606060',
            'selection': '#0078d4',
            'shadow': '#000000',
            'gradient_start': '#2d2d2d',
            'gradient_end': '#1a1a1a'
        },
        'pop_bright': {  # POPで明るい
            'name': '🌈 Pop Bright',
            'bg_primary': '#F8F9FA',
            'bg_secondary': '#FFFFFF',
            'bg_tertiary': '#F1F3F4',
            'bg_quaternary': '#E8EAED',
            'bg_accent': '#6C5CE7',
            'bg_hover': '#5B4BD3',
            'bg_pressed': '#4A3BC1',
            'fg_primary': '#2D3436',
            'fg_secondary': '#636E72',
            'fg_tertiary': '#B2BEC3',
            'fg_accent': '#FD79A8',
            'success': '#00B894',
            'warning': '#FDCB6E',
            'error': '#E84393',
            'info': '#00CEC9',
            'border': '#DDD6FE',
            'border_light': '#E5E5E5',
            'selection': '#6C5CE7',
            'shadow': '#00000020',
            'gradient_start': '#FFFFFF',
            'gradient_end': '#F8F9FA'
        },
        'cyber_neon': {  # サイバーネオン
            'name': '⚡ Cyber Neon',
            'bg_primary': '#0A0A0A',
            'bg_secondary': '#1A1A2E',
            'bg_tertiary': '#16213E',
            'bg_quaternary': '#0F3460',
            'bg_accent': '#00FF88',
            'bg_hover': '#00DD77',
            'bg_pressed': '#00BB66',
            'fg_primary': '#FFFFFF',
            'fg_secondary': '#00FFFF',
            'fg_tertiary': '#888888',
            'fg_accent': '#FF0080',
            'success': '#00FF88',
            'warning': '#FFD700',
            'error': '#FF4081',
            'info': '#00BFFF',
            'border': '#00FF8850',
            'border_light': '#00FFFF30',
            'selection': '#00FF88',
            'shadow': '#00FF8830',
            'gradient_start': '#1A1A2E',
            'gradient_end': '#16213E'
        }
    }
    
    # デフォルトテーマ
    current_theme = 'dark_pro'
    THEME = THEMES[current_theme]
    
    # 日本語美しいフォント設定（太字潰れ対策済み）
    FONTS = {
        'title': ('Meiryo UI', 18, 'normal'),      # メインタイトル（太字なし）
        'subtitle': ('Meiryo UI', 13, 'normal'),   # サブタイトル（太字なし）
        'heading': ('Meiryo UI', 11, 'normal'),    # 見出し（太字なし）
        'body': ('Meiryo UI', 10),                 # 本文
        'small': ('Meiryo UI', 9),                 # 小さいテキスト
        'button': ('Meiryo UI', 9, 'normal'),      # ボタン（太字なし）
        'tree': ('Meiryo UI', 9),                  # ツリー
        'status': ('Meiryo UI', 9),                # ステータスバー
        'tree_header': ('Meiryo UI', 10, 'normal'), # ツリーヘッダー（太字なし）
        'tooltip': ('Meiryo UI', 8),               # ツールチップ
        'tiny': ('Meiryo UI', 8),                  # 極小テキスト
        'code': ('Consolas', 9),                   # コード表示
        'japanese': ('Meiryo UI', 10)              # 日本語専用
    }
    
    
    # キーオプション
    KEY_OPTIONS = [
        "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12",
        "enter", "tab", "space", "backspace", "delete", "esc", "home", "end", "pageup", "pagedown", "up", "down", "left", "right",
        "ctrl+c", "ctrl+v", "ctrl+a", "ctrl+z", "ctrl+y", "ctrl+x", "ctrl+s", "ctrl+o", "ctrl+n", "ctrl+f", "ctrl+h", "ctrl+r",
        "shift+tab", "shift+f10", "shift+delete", "shift+insert",
        "alt+tab", "alt+F4", "alt+left", "alt+right", "alt+enter",
        "win+d", "win+e", "win+r", "win+l", "win+tab",
        "ctrl+alt+delete", "ctrl+shift+esc", "ctrl+shift+n", "ctrl+shift+t",
        "ctrl+shift+f12",
    ]
    
    CLICK_TYPES = ["single", "double", "right"]
    
    # Treeview拡張機能の定数
    FG_DISABLED = '#8E8E8E'  # 無効行の文字色
    ELLIPSIS = '…'  # 省略記号
    TOOLTIP_MAX_WIDTH_PX = 640  # ツールチップの最大幅
    
    # レイアウト調整の定数
    SPACING_SCALE = 1.15  # 余白のスケール調整（1.0/1.15/1.3で切り替え可能）
    
    # 🎯 統一アイコンセット
    STEP_ICONS = {
        'image_click': '🖱️',
        'coord_click': '📍', 
        'coord_drag': '↔️',
        'image_relative_right_click': '🎯',
        'sleep': '⏱️',
        'key': '⌨️',
        'copy': '📋',
        'paste': '📄',
        'custom_text': '📝',
        'cmd_command': '💻',
        'repeat_start': '🔁',
        'repeat_end': '🔚',
        'screenshot': '📷'
    }
    
    # 🎨 カテゴリ別カラーアイコン
    CATEGORY_COLORS = {
        'click': '#E74C3C',    # 赤系
        'input': '#3498DB',    # 青系  
        'wait': '#F39C12',     # オレンジ系
        'system': '#9B59B6',   # 紫系
        'repeat': '#27AE60'    # 緑系
    }
    
    @classmethod
    def ensure_directories(cls):
        """必要なディレクトリを作成"""
        dirs = ['config', 'logs', 'images', 'backups', 'temp']
        for dir_name in dirs:
            Path(dir_name).mkdir(exist_ok=True)
    
    
    
    @classmethod
    def get_step_icon(cls, step_type: str):
        """ステップタイプに対応するアイコンを取得"""
        return cls.STEP_ICONS.get(step_type, '⚙️')
    
    @classmethod
    def get_step_display_name(cls, step_type: str):
        """ステップタイプの表示名を取得（テキストのみ）"""
        display_names = {
            'image_click': '画像クリック',
            'coord_click': '座標クリック', 
            'coord_drag': '座標ドラッグ',
            'image_relative_right_click': '画像オフセット',
            'sleep': '待機',
            'key': 'キー入力',
            'copy': 'コピー',
            'paste': '貼り付け',
            'custom_text': 'テキスト入力',
            'cmd_command': 'コマンド実行',
            'repeat_start': '繰り返し開始',
            'repeat_end': '繰り返し終了',
            'screenshot': 'スクリーンショット'
        }
        return display_names.get(step_type, step_type)
    
    @classmethod
    def apply_dark_theme(cls, root):
        """次世代プロフェッショナルテーマを適用"""
        style = ttk.Style()
        
        # テーマを設定
        style.theme_use('clam')
        
        # メインコンテナスタイル
        style.configure('Modern.TFrame', background=cls.THEME['bg_secondary'], relief='flat')
        style.configure('Card.TFrame', background=cls.THEME['bg_tertiary'], relief='raised', borderwidth=1)
        style.configure('Toolbar.TFrame', background=cls.THEME['bg_quaternary'], relief='flat')
        
        # 🎴 モダンカードスタイル
        style.configure('ModernCard.TFrame', 
                       background=cls.THEME['bg_tertiary'], 
                       relief='solid', 
                       borderwidth=1)
        style.configure('HighlightCard.TFrame', 
                       background=cls.THEME['bg_accent'], 
                       relief='solid', 
                       borderwidth=2)
        style.configure('CompactCard.TFrame', 
                       background=cls.THEME['bg_quaternary'], 
                       relief='groove', 
                       borderwidth=1)
        
        # ラベルスタイル群
        style.configure('Modern.TLabel', background=cls.THEME['bg_secondary'], 
                       foreground=cls.THEME['fg_primary'], font=cls.FONTS['body'])
        style.configure('Title.TLabel', background=cls.THEME['bg_secondary'], 
                       foreground=cls.THEME['fg_accent'], font=cls.FONTS['title'])
        style.configure('Subtitle.TLabel', background=cls.THEME['bg_secondary'], 
                       foreground=cls.THEME['fg_primary'], font=cls.FONTS['subtitle'])
        style.configure('Heading.TLabel', background=cls.THEME['bg_secondary'], 
                       foreground=cls.THEME['fg_secondary'], font=cls.FONTS['heading'])
        style.configure('Caption.TLabel', background=cls.THEME['bg_secondary'], 
                       foreground=cls.THEME['fg_tertiary'], font=cls.FONTS['small'])
        
        # LabelFrameスタイル
        style.configure('Modern.TLabelframe', background=cls.THEME['bg_secondary'], 
                       foreground=cls.THEME['fg_primary'], borderwidth=1, relief='solid')
        style.configure('Modern.TLabelframe.Label', background=cls.THEME['bg_secondary'], 
                       foreground=cls.THEME['fg_primary'], font=cls.FONTS['body'])
        
        # プロフェッショナルボタンスタイル - 視覚調整
        dialog_button_font = ('Meiryo UI', 10, 'normal')  # フォント統一
        style.configure('Modern.TButton', 
                       background=cls.THEME['bg_accent'],
                       foreground=cls.THEME['fg_primary'],
                       font=dialog_button_font,  # 統一フォント
                       borderwidth=0,
                       focuscolor='none',
                       relief='flat',
                       padding=(16, 14),  # キャンセルボタンを視覚的に大きく
                       width=10)  # 最小幅を統一
        style.map('Modern.TButton',
                  background=[('active', cls.THEME['bg_hover']),
                             ('pressed', cls.THEME['bg_pressed']),
                             ('disabled', cls.THEME['bg_tertiary'])],
                  foreground=[('disabled', cls.THEME['fg_tertiary'])])
        
        # コンパクトボタンスタイル（編集操作用）
        style.configure('Compact.TButton', 
                       background=cls.THEME['bg_accent'],
                       foreground=cls.THEME['fg_primary'],
                       font=('Meiryo UI', 9, 'normal'),  # 他に合わせて太字解除、サイズ調整
                       borderwidth=0,
                       focuscolor='none',
                       relief='flat',
                       padding=(10, 6),  # padding調整（8,6→10,6）横幅少し大きく
                       width=4)  # width 3→4に少し拡大
        style.map('Compact.TButton',
                  background=[('active', cls.THEME['bg_hover']),
                             ('pressed', cls.THEME['bg_pressed']),
                             ('disabled', cls.THEME['bg_tertiary'])],
                  foreground=[('disabled', cls.THEME['fg_tertiary'])])
        
        # 特別ボタンスタイル - 視覚調整
        style.configure('Primary.TButton', 
                       background=cls.THEME['success'],
                       foreground=cls.THEME['fg_primary'],
                       font=dialog_button_font,  # 統一フォント
                       borderwidth=0,
                       relief='flat',
                       padding=(16, 14),  # 両ボタンを同じサイズに統一
                       width=10)  # 最小幅を統一
        style.map('Primary.TButton',
                  background=[('active', '#14b047'), ('pressed', '#0e8c38')])
        
        style.configure('Danger.TButton', 
                       background=cls.THEME['error'],
                       foreground=cls.THEME['fg_primary'],
                       font=dialog_button_font,  # Primary.TButtonと同じフォントに統一
                       borderwidth=0,
                       relief='flat',
                       padding=(16, 14))  # Primary.TButtonと同じpaddingに統一
        style.map('Danger.TButton',
                  background=[('active', '#c0392b'), ('pressed', '#a93226')])
        
        # 🎨 視認性の高いカテゴリ別ボタンカラー
        # 🎨 意味のある色の配色 - 直感的に理解しやすい色使い
        category_styles = [
            ('Click.TButton', '#1E88E5', '#1565C0'),      # 🔵 信頼できるブルー（クリック操作）
            ('Input.TButton', '#43A047', '#2E7D32'),      # 🟢 アクティブグリーン（入力操作）  
            ('Wait.TButton', '#FB8C00', '#EF6C00'),       # 🟡 注意喚起アンバー（待機操作）
            ('System.TButton', '#8E24AA', '#6A1B9A'),     # 🟣 高度機能パープル（システム操作）
            ('Repeat.TButton', '#444444', '#2A2A2A')      # 🔘 黒寄りの濃い灰色（繰り返し操作）

        ]
        
        for style_name, bg_color, hover_color in category_styles:
            style.configure(style_name,
                           background=bg_color,
                           foreground='#FFFFFF',
                           font=cls.FONTS['button'],
                           borderwidth=0,
                           relief='flat',
                           padding=(12, 8))
            style.map(style_name,
                     background=[('active', hover_color),
                                ('pressed', hover_color)],
                     foreground=[('active', '#FFFFFF'),
                                ('pressed', '#FFFFFF')])
                                
        # 🌟 特別なアクセントボタン
        style.configure('Accent.TButton',
                       background=cls.THEME['fg_accent'],
                       foreground=cls.THEME['bg_primary'],
                       font=cls.FONTS['button'],
                       borderwidth=0,
                       relief='flat',
                       padding=(15, 10))
        style.map('Accent.TButton',
                  background=[('active', '#4ABFCD'), ('pressed', '#3A9BA8')])

        # 🎯 大きな実行・停止ボタン専用スタイル
        style.configure('Large.Primary.TButton',
                       background='#4CAF50',
                       foreground='#FFFFFF',
                       font=('Meiryo UI', 14, 'normal'),
                       borderwidth=0,
                       relief='flat',
                       padding=(30, 20))  # 大きなpadding
        style.map('Large.Primary.TButton',
                  background=[('active', '#45A049'), ('pressed', '#3E8E41')])

        style.configure('Large.Danger.TButton',
                       background='#F44336',
                       foreground='#FFFFFF',
                       font=('Meiryo UI', 14, 'normal'),
                       borderwidth=0,
                       relief='flat',
                       padding=(30, 20))  # 大きなpadding
        style.map('Large.Danger.TButton',
                  background=[('active', '#E53935'), ('pressed', '#C62828')])
        
        # 高級ツリービュースタイル（境界線強化）
        # Modern系スタイル
        style.configure('Modern.Treeview', 
                       background=cls.THEME['bg_secondary'],
                       foreground=cls.THEME['fg_primary'],
                       fieldbackground=cls.THEME['bg_secondary'],
                       borderwidth=1,
                       relief='solid',
                       font=cls.FONTS['tree'],
                       rowheight=26)  # 少し縮小してコンパクトに
        style.configure('Modern.Treeview.Heading',
                       background=cls.THEME['bg_quaternary'],
                       foreground=cls.THEME['fg_primary'],
                       font=cls.FONTS['tree_header'],
                       relief='raised',
                       borderwidth=1)
        style.map('Modern.Treeview',
                  background=[('selected', cls.THEME['selection']),
                              ('active', cls.THEME['bg_hover'])],
                  foreground=[('selected', cls.THEME['fg_primary'])])

        # 互換性のためデフォルトのTreeview/Headingにも適用（環境によってModern.*が反映されない場合の保険）
        style.configure('Treeview',
                        background=cls.THEME['bg_secondary'],
                        foreground=cls.THEME['fg_primary'],
                        fieldbackground=cls.THEME['bg_secondary'],
                        borderwidth=1,
                        relief='solid',
                        font=cls.FONTS['tree'],
                        rowheight=26)
        style.configure('Treeview.Heading',
                        background=cls.THEME['bg_quaternary'],
                        foreground=cls.THEME['fg_primary'],
                        font=cls.FONTS['tree_header'],
                        relief='raised',
                        borderwidth=1)
        
        # 入力フィールドスタイル
        style.configure('Modern.TEntry',
                       fieldbackground=cls.THEME['bg_tertiary'],
                       foreground=cls.THEME['fg_primary'],
                       borderwidth=1,
                       relief='solid',
                       bordercolor=cls.THEME['border'],
                       lightcolor=cls.THEME['border_light'],
                       darkcolor=cls.THEME['border'],
                       focuscolor=cls.THEME['bg_accent'],
                       insertcolor=cls.THEME['fg_primary'])
        
        # Comboboxスタイル
        style.map('TCombobox',
                  fieldbackground=[('readonly', cls.THEME['bg_tertiary']), ('!readonly', cls.THEME['bg_tertiary'])],
                  background=[('readonly', cls.THEME['bg_tertiary']), ('!readonly', cls.THEME['bg_tertiary'])],
                  foreground=[('readonly', cls.THEME['fg_primary']), ('!readonly', cls.THEME['fg_primary'])],
                  selectbackground=[('readonly', cls.THEME['bg_tertiary']), ('!readonly', cls.THEME['bg_tertiary'])],
                  selectforeground=[('readonly', cls.THEME['fg_primary']), ('!readonly', cls.THEME['fg_primary'])],
                  arrowcolor=[('readonly', cls.THEME['fg_primary']), ('!readonly', cls.THEME['fg_primary'])],
                  bordercolor=[('readonly', cls.THEME['border']), ('!readonly', cls.THEME['border'])])
        style.configure('TCombobox',
                        arrowsize=20,
                        relief='flat')

        # Checkbuttonスタイル
        style.configure('Modern.TCheckbutton',
                        background=cls.THEME['bg_secondary'],
                        foreground=cls.THEME['fg_primary'],
                        indicatorcolor=cls.THEME['bg_tertiary'],
                        font=cls.FONTS['body'])
        style.map('Modern.TCheckbutton',
                  indicatorcolor=[('selected', cls.THEME['bg_accent']),
                                  ('active', cls.THEME['bg_hover'])])
        
        # スクロールバースタイル
        style.configure('Modern.Vertical.TScrollbar',
                       background=cls.THEME['bg_quaternary'],
                       darkcolor=cls.THEME['bg_tertiary'],
                       lightcolor=cls.THEME['bg_quaternary'],
                       troughcolor=cls.THEME['bg_secondary'],
                       borderwidth=0,
                       arrowcolor=cls.THEME['fg_secondary'])
        
        style.configure('Modern.Horizontal.TScrollbar',
                       background=cls.THEME['bg_quaternary'],
                       darkcolor=cls.THEME['bg_tertiary'],
                       lightcolor=cls.THEME['bg_quaternary'],
                       troughcolor=cls.THEME['bg_secondary'],
                       borderwidth=0,
                       arrowcolor=cls.THEME['fg_secondary'])
        
        # PanedWindowスタイル
        style.configure('Modern.TPanedWindow', 
                       background=cls.THEME['bg_secondary'],
                       borderwidth=1,
                       relief='flat')
        
        # ルートウィンドウ設定
        root.configure(bg=cls.THEME['bg_primary'])


class BaseDialog:
    """共通のダイアログ機能を提供する基底クラス"""
    
    def __init__(self, parent: tk.Tk, title: str, width: int = 700, height: int = 600):
        self.parent = parent
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry(f"{width}x{height}")
        self.dialog.configure(bg=AppConfig.THEME['bg_primary'])
        self.dialog.resizable(True, True)
        self.dialog.minsize(500, 400)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.center_window()
        
        # ESCキーでダイアログを閉じる
        self.dialog.bind('<Escape>', lambda e: self.dialog.destroy())
    
    def center_window(self):
        """ウィンドウをスクリーン中央に配置"""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f'+{x}+{y}')
    
    def show_error_dialog(self, title: str, message: str, parent=None):
        """エラーダイアログを表示"""
        if parent is None:
            parent = self.dialog
        messagebox.showerror(title, message, parent=parent)
    
    def show_info_dialog(self, title: str, message: str, parent=None):
        """情報ダイアログを表示"""
        if parent is None:
            parent = self.dialog
        messagebox.showinfo(title, message, parent=parent)
    
    def show_warning_dialog(self, title: str, message: str, parent=None):
        """警告ダイアログを表示"""
        if parent is None:
            parent = self.dialog
        messagebox.showwarning(title, message, parent=parent)


class MessageBoxUtils:
    """メッセージボックスのユーティリティクラス"""
    
    @staticmethod
    def show_error(title: str, message: str, parent=None):
        """エラーダイアログを表示"""
        messagebox.showerror(title, message, parent=parent)
    
    @staticmethod
    def show_info(title: str, message: str, parent=None):
        """情報ダイアログを表示"""
        messagebox.showinfo(title, message, parent=parent)
    
    @staticmethod
    def show_warning(title: str, message: str, parent=None):
        """警告ダイアログを表示"""
        messagebox.showwarning(title, message, parent=parent)
    
    @staticmethod
    def ask_yes_no(title: str, message: str, parent=None) -> bool:
        """はい/いいえダイアログを表示"""
        return messagebox.askyesno(title, message, parent=parent)
    
    @staticmethod
    def ask_ok_cancel(title: str, message: str, parent=None) -> bool:
        """OK/キャンセルダイアログを表示"""
        return messagebox.askokcancel(title, message, parent=parent)


@dataclass
class Step:
    """実行ステップを表すデータクラス"""
    type: str
    params: Dict[str, Any]
    comment: str
    created_at: str = ""
    enabled: bool = True
    
    def __post_init__(self):
        if not self.created_at:
            self.created_at = datetime.now().isoformat()
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Step':
        # 古いフォーマットとの互換性を保つ
        if 'enabled' not in data:
            data['enabled'] = True
        return cls(**data)
    
    def validate(self) -> bool:
        """ステップの妥当性をチェック"""
        required_fields = {'type', 'params', 'comment'}
        return all(hasattr(self, field) for field in required_fields)
    
    def get_preview_image_path(self) -> Optional[str]:
        """プレビュー用の画像パスを取得"""
        if self.type in ['image_click', 'image_relative_right_click']:
            return self.params.get('path')
        return None

class ImagePreviewWidget:
    """画像プレビューウィジェット"""
    
    def __init__(self, parent, app_instance=None):
        self.parent = parent
        self.app_instance = app_instance
        self.frame = ttk.Frame(parent, style='Modern.TFrame')
        self.current_image_path = None
        self.image_label = None
        self.info_label = None
        self.setup_ui()
    
    def setup_ui(self):
        # シンプルなプレビューのみ（タブなし）
        self.preview_frame = ttk.Frame(self.frame)
        
        self.preview_frame.pack(fill='both', expand=True)
        self.setup_preview_only()
        
        # 情報表示エリア（幅固定）
        self.info_label = ttk.Label(self.frame, text="", 
                                   style='Modern.TLabel',
                                   wraplength=200,  # 最大幅を200pxに制限
                                   justify='center')
        self.info_label.pack()

    def setup_preview_only(self):
        """シンプルなプレビューのみのUI構築"""
        # 画像表示エリア
        self.image_frame = tk.Frame(self.preview_frame, 
                                   bg=AppConfig.THEME['bg_tertiary'],
                                   relief='sunken', bd=2)
        self.image_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # プレースホルダー
        self.placeholder_label = ttk.Label(self.image_frame, 
                                          text="画像ステップを選択すると\nプレビューが表示されます",
                                          style='Modern.TLabel',
                                          anchor='center')
        self.placeholder_label.pack(expand=True)
    
    def show_image(self, image_path: str):
        """画像を表示"""
        try:
            if not image_path or not Path(image_path).exists():
                self.clear_image()
                return
            
            # 画像を読み込み
            pil_image = Image.open(image_path)
            
            # サイズ調整（コンパクトサイズ）
            display_size = (160, 90)  # 高さをさらに縮小してコンパクトに
            pil_image.thumbnail(display_size, Image.Resampling.LANCZOS)
            
            # Tkinter用に変換
            photo = ImageTk.PhotoImage(pil_image)
            
            # プレースホルダーを非表示
            self.placeholder_label.pack_forget()
            
            # 画像ラベルを作成/更新
            if self.image_label is None:
                self.image_label = tk.Label(self.image_frame, 
                                           bg=AppConfig.THEME['bg_tertiary'])
                self.image_label.pack(expand=True)
            
            self.image_label.configure(image=photo)
            self.image_label.image = photo  # 参照を保持
            
            # 情報表示を更新（ファイル名を省略形で表示）
            file_size = Path(image_path).stat().st_size
            original_size = Image.open(image_path).size
            filename = Path(image_path).name
            # ファイル名が長い場合は省略
            if len(filename) > 20:
                filename = filename[:15] + "..." + filename[-5:]
            info_text = f"{filename} | {original_size[0]}x{original_size[1]}px | {file_size:,}B"
            self.info_label.configure(text=info_text)
            
            self.current_image_path = image_path
            
            # 一致テストを更新
            if hasattr(self, 'threshold_var'):
                self.update_match_test()
            
        except Exception as e:
            logger.error(f"画像表示エラー: {e}")
            self.clear_image()

    
    def clear_image(self):
        """画像をクリア"""
        if self.image_label:
            self.image_label.pack_forget()
            self.image_label = None
        
        self.placeholder_label.pack(expand=True)
        self.info_label.configure(text="")
        self.current_image_path = None

class DragDropTreeview:
    """ドラッグ&ドラップ対応のTreeview"""
    
    def __init__(self, parent, app_instance):
        self.app = app_instance
        self.drag_item = None
        self.drag_y = 0
        
        # ドラッグ用フローティング表示
        self.drag_window = None
        self.original_values = None
        
        # 省略表示とツールチップ機能用の変数
        self.full_values = {}  # {item_id: {'Params': '完全なテキスト', 'Comment': '完全なテキスト'}}
        self.tooltip_window = None
        self.current_tooltip_item = None
        self.current_tooltip_column = None
        
        # Treeviewを作成
        self.tree = ttk.Treeview(parent, 
                                height=15, # 表示行数を指定
                                columns=("Status", "Line", "Type", "Params", "Comment"), 
                                show="headings",
                                style='Modern.Treeview')
        
        # カラム設定（境界線を明確化）
        self.tree.heading("Status", text="")
        self.tree.heading("Line", text="#")
        self.tree.heading("Type", text="アクション")
        self.tree.heading("Params", text="内容") 
        self.tree.heading("Comment", text="メモ")
        
        # 最適化されたカラム設定（メモ欄を広く）
        self.tree.column("Status", minwidth=30, width=35, stretch=False)
        self.tree.column("Line", minwidth=30, width=35, stretch=False)  
        self.tree.column("Type", minwidth=100, width=120, stretch=False)      
        self.tree.column("Params", minwidth=120, width=150, stretch=True)     
        self.tree.column("Comment", minwidth=200, width=300, stretch=False)
        
        # ドラッグ&ドラップイベントをバインド
        self.tree.bind("<Button-1>", self.on_drag_start)
        self.tree.bind("<B1-Motion>", self.on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self.on_drag_release)
        
        # 無効行のタグ設定
        self.tree.tag_configure('disabled', foreground=AppConfig.FG_DISABLED)
        # ドラッグ中のタグ設定（フォントサイズも大きく）
        drag_font = (AppConfig.FONTS['tree'][0], AppConfig.FONTS['tree'][1] + 2, 'bold')
        self.tree.tag_configure('dragging', background='#4A90E2', foreground='white', font=drag_font)
        
        # 🎯 強化: ファイルドロップ対応
        self.setup_file_drop()
        
        # その他のイベント
        self.tree.bind("<ButtonRelease-1>", self.on_select, add='+')
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-3>", self.on_right_click)
        self.tree.bind("<Return>", self.on_enter)
        self.tree.bind("<Delete>", self.on_delete)
        self.tree.bind("<space>", self.on_toggle_enabled)
        
        # ツールチップ用のイベント
        self.tree.bind("<Motion>", self.on_motion)
        self.tree.bind("<Leave>", self.on_leave)
        
        # 列幅変更時の再計算
        self.tree.bind("<ButtonRelease-1>", self.on_column_resize, add='+')

        # macOS等でheadingsのみだと表示が不安定になるケースに備え、#0を明示的に隠す
        try:
            self.tree.column('#0', width=0, stretch=False)
        except Exception:
            pass
    
    def create_drag_window(self, item):
        """ドラッグ用のフローティングウィンドウを作成"""
        if self.drag_window:
            self.drag_window.destroy()
        
        # フローティングウィンドウを作成
        self.drag_window = tk.Toplevel(self.tree)
        self.drag_window.wm_overrideredirect(True)  # ボーダーなし
        self.drag_window.configure(bg='#4A90E2')
        self.drag_window.attributes('-alpha', 0.8)  # 半透明
        self.drag_window.attributes('-topmost', True)  # 最前面
        
        # 元の行の値を取得
        values = self.tree.item(item, 'values')
        self.original_values = values
        
        # フレームを作成（パディング用）
        frame = tk.Frame(self.drag_window, bg='#4A90E2', padx=5, pady=3)
        frame.pack(fill='both', expand=True)
        
        # 行の内容を表示（簡略化）
        display_text = f"🔄 {values[2]} | {values[3][:30]}{'...' if len(values[3]) > 30 else ''}"
        label = tk.Label(frame, text=display_text, 
                        bg='#4A90E2', fg='white',
                        font=(AppConfig.FONTS['tree'][0], AppConfig.FONTS['tree'][1] + 1, 'bold'))
        label.pack()
        
        # ウィンドウを非表示にしておく（位置調整後に表示）
        self.drag_window.withdraw()
    
    def update_drag_window_position(self, event):
        """フローティングウィンドウの位置を更新"""
        if self.drag_window:
            # マウス位置にウィンドウを配置（少し下にずらす）
            x = event.x_root + 10
            y = event.y_root + 10
            self.drag_window.geometry(f"+{x}+{y}")
            if self.drag_window.state() == 'withdrawn':
                self.drag_window.deiconify()
    
    def hide_drag_window(self):
        """フローティングウィンドウを非表示"""
        if self.drag_window:
            self.drag_window.destroy()
            self.drag_window = None
            self.original_values = None
    
    def on_drag_start(self, event):
        """ドラッグ開始準備（実際のドラッグはモーション時に判定）"""
        item = self.tree.identify_row(event.y)
        if item:
            self.drag_item = item
            self.drag_y = event.y
            # フローティングウィンドウはまだ作成しない（motion時に作成）
    
    def on_drag_motion(self, event):
        """ドラッグ中"""
        if self.drag_item and abs(event.y - self.drag_y) > 20:
            # 実際のドラッグ開始時にフローティングウィンドウを作成
            if not self.drag_window:
                self.create_drag_window(self.drag_item)
            
            # フローティングウィンドウの位置を更新
            self.update_drag_window_position(event)
            
            # ドラッグ視覚効果を適用
            current_values = list(self.tree.item(self.drag_item, 'values'))
            if not current_values[0].startswith('🔄'):
                current_values[0] = '🔄'  # 行番号に🔄アイコン
                self.tree.item(self.drag_item, values=current_values, tags=('dragging',))
            
            # ドラッグインジケータを表示（簡易版）
            target_item = self.tree.identify_row(event.y)
            if target_item and target_item != self.drag_item:
                self.tree.selection_set(target_item)
    
    def on_drag_release(self, event):
        """ドラッグ終了"""
        if not self.drag_item:
            return
        
        # ドラッグ距離が短い場合（シングルクリック）の処理
        drag_distance = abs(event.y - self.drag_y)
        is_single_click = drag_distance <= 20
        
        # フローティングウィンドウを非表示
        self.hide_drag_window()
        
        # ドラッグ視覚効果をリセット
        current_values = list(self.tree.item(self.drag_item, 'values'))
        if current_values[0] == '🔄':
            drag_index = self.tree.index(self.drag_item)
            current_values[0] = str(drag_index + 1)  # 元の行番号に戻す
            self.tree.item(self.drag_item, values=current_values, tags=())
        
        if is_single_click:
            # シングルクリックの場合は通常の選択処理
            self.tree.selection_set(self.drag_item)
            self.tree.focus(self.drag_item)
            # 選択イベントを手動で呼び出し
            self.on_select_manual(self.drag_item)
        else:
            # ドラッグによる移動処理
            target_item = self.tree.identify_row(event.y)
            if target_item and target_item != self.drag_item:
                # アイテムの移動を実行
                drag_index = self.tree.index(self.drag_item)
                target_index = self.tree.index(target_item)
                
                # ステップの順序を変更
                self.app.move_step(drag_index, target_index)
            else:
                # 移動しなかった場合も視覚効果をリセット
                self.app.refresh_tree()
        
        self.drag_item = None
        self.drag_y = 0
    
    def on_select_manual(self, item):
        """手動選択処理（ドラッグ終了時のシングルクリック用）"""
        if item and hasattr(self.app, 'image_preview'):
            index = self.tree.index(item)
            if index < len(self.app.steps):
                step = self.app.steps[index]
                image_path = step.get_preview_image_path()
                if image_path:
                    self.app.image_preview.show_image(image_path)
                else:
                    self.app.image_preview.clear_image()
    
    def on_select(self, event):
        """選択イベント"""
        selection = self.tree.selection()
        if selection and hasattr(self.app, 'image_preview'):
            index = self.tree.index(selection[0])
            if index < len(self.app.steps):
                step = self.app.steps[index]
                image_path = step.get_preview_image_path()
                if image_path:
                    self.app.image_preview.show_image(image_path)
                else:
                    self.app.image_preview.clear_image()
    
    def on_double_click(self, event):
        """ダブルクリックイベント"""
        self.app.edit_selected_step()
    
    def on_right_click(self, event):
        """右クリックイベント"""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.show_context_menu(event)
    
    def on_enter(self, event):
        """エンターキーイベント"""
        self.app.edit_selected_step()
    
    def on_delete(self, event):
        """削除キーイベント"""
        self.app.delete_selected()
    
    def on_toggle_enabled(self, event):
        """スペースキーで有効/無効切り替え"""
        selection = self.tree.selection()
        if selection:
            index = self.tree.index(selection[0])
            if index < len(self.app.steps):
                self.app.toggle_step_enabled(index)
    
    def apply_row_state_tags(self, item_id, enabled: bool):
        """行の有効/無効状態に応じてタグを適用"""
        if enabled:
            self.tree.item(item_id, tags=())  # タグを削除
        else:
            self.tree.item(item_id, tags=('disabled',))  # 無効タグを適用
    
    def elide_to_fit(self, text: str, column: str) -> str:
        """列幅に収まるようにテキストを省略"""
        import tkinter.font as tkfont
        
        # フォントとカラム幅を取得（TreeviewのデフォルトフォントはTkDefaultFontを使用）
        try:
            font = tkfont.nametofont("TkDefaultFont")
        except:
            font = tkfont.Font(family="Segoe UI", size=9)
        column_width = self.tree.column(column, 'width')
        
        # テキストの幅を測定
        text_width = font.measure(text)
        
        if text_width <= column_width - 10:  # 10pxのマージンを考慮
            return text
        
        # 省略が必要な場合
        ellipsis_width = font.measure(AppConfig.ELLIPSIS)
        available_width = column_width - ellipsis_width - 10
        
        # 文字を一つずつ削って、幅に収まるまで縮める
        for i in range(len(text), 0, -1):
            truncated = text[:i]
            if font.measure(truncated) <= available_width:
                return truncated + AppConfig.ELLIPSIS
        
        return AppConfig.ELLIPSIS
    
    def on_motion(self, event):
        """マウス移動時のツールチップ処理"""
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        
        if not item or not column:
            self.hide_tooltip()
            return
        
        # カラム名を取得（#1, #2, #3...を実際の名前に変換）
        column_names = ["Status", "Line", "Type", "Params", "Comment"]
        try:
            col_index = int(column.replace('#', '')) - 1
            if 0 <= col_index < len(column_names):
                column_name = column_names[col_index]
            else:
                return
        except:
            return
        
        # 省略表示対象の列のみ処理（ParamsとComment）
        if column_name not in ["Params", "Comment"]:
            self.hide_tooltip()
            return
        
        # 同じセルの場合は何もしない
        if self.current_tooltip_item == item and self.current_tooltip_column == column_name:
            return
        
        # ツールチップを非表示
        self.hide_tooltip()
        
        # 完全なテキストがある場合のみツールチップを表示
        if item in self.full_values and column_name in self.full_values[item]:
            full_text = self.full_values[item][column_name]
            displayed_text = self.tree.set(item, column_name)
            
            # 省略されている場合のみツールチップを表示
            if displayed_text.endswith(AppConfig.ELLIPSIS):
                self.show_tooltip(event, full_text)
                self.current_tooltip_item = item
                self.current_tooltip_column = column_name
    
    def on_leave(self, event):
        """マウスがTreeviewから離れた時"""
        self.hide_tooltip()
    
    def show_tooltip(self, event, text: str):
        """ツールチップを表示"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
        
        self.tooltip_window = tw = tk.Toplevel(self.tree)
        tw.wm_overrideredirect(True)
        
        # ツールチップの位置を計算
        x = event.x_root + 10
        y = event.y_root + 10
        
        # スクリーンの端を考慮して位置を調整
        screen_width = tw.winfo_screenwidth()
        screen_height = tw.winfo_screenheight()
        
        # ツールチップの内容を作成
        import tkinter.font as tkfont
        font = tkfont.Font(family="Meiryo UI", size=9)
        
        # 最大幅を考慮してテキストを改行
        lines = []
        words = text.split()
        current_line = ""
        
        for word in words:
            test_line = current_line + (" " + word if current_line else word)
            if font.measure(test_line) <= AppConfig.TOOLTIP_MAX_WIDTH_PX:
                current_line = test_line
            else:
                if current_line:
                    lines.append(current_line)
                current_line = word
        
        if current_line:
            lines.append(current_line)
        
        final_text = "\n".join(lines)
        
        label = tk.Label(tw, text=final_text, font=font,
                        background="#FFFFE0", foreground="black",
                        borderwidth=1, relief="solid", padx=6, pady=4)
        label.pack()
        
        # ツールチップのサイズを取得して位置を調整
        tw.update_idletasks()
        tooltip_width = tw.winfo_reqwidth()
        tooltip_height = tw.winfo_reqheight()
        
        # スクリーンからはみ出ないように調整
        if x + tooltip_width > screen_width:
            x = screen_width - tooltip_width - 10
        if y + tooltip_height > screen_height:
            y = event.y_root - tooltip_height - 10
        
        tw.geometry(f"+{x}+{y}")
    
    def hide_tooltip(self):
        """ツールチップを非表示"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
        self.current_tooltip_item = None
        self.current_tooltip_column = None
    
    def on_column_resize(self, event):
        """列幅変更時の再計算"""
        # 列幅変更の可能性がある場合のみ処理
        region = self.tree.identify_region(event.x, event.y)
        if region == "separator":
            # 少し遅延してから再計算（列幅変更が完了してから）
            self.tree.after(10, self.recalculate_display)
    
    def recalculate_display(self):
        """省略表示の再計算"""
        try:
            for item_id in self.tree.get_children():
                if item_id in self.full_values:
                    # ParamsとCommentの表示を再計算
                    for column_name in ['Params', 'Comment']:
                        if column_name in self.full_values[item_id]:
                            full_text = self.full_values[item_id][column_name]
                            display_text = self.elide_to_fit(full_text, column_name)
                            self.tree.set(item_id, column_name, display_text)
        except Exception as e:
            # エラーが発生しても処理を続行
            pass
    
    def setup_file_drop(self):
        """🎯 ファイルドロップ機能設定"""
        try:
            # tkinterdnd2ライブラリを使用してファイルドロップを設定
            # ライブラリが利用できない場合は静かに機能を無効化
            
            # まず、必要なメソッドが存在するかチェック
            if not hasattr(self.tree, 'drop_target_register'):
                # tkinterdnd2がインストールされていない場合は何もしない
                # ログも出力しない（ユーザーに混乱を与えないため）
                return
            
            def on_drop(event):
                """ファイルドロップ時の処理"""
                try:
                    files = self.tree.tk.splitlist(event.data)
                    for file_path in files:
                        self.handle_dropped_file(file_path)
                except Exception as e:
                    logger.error(f"ファイルドロップエラー: {e}")
            
            # ドロップ受け入れ設定
            self.tree.drop_target_register('DND_Files')
            self.tree.dnd_bind('<<Drop>>', on_drop)
            
            # ファイルドロップ機能が正常に設定された場合のみログ出力
            logger.debug("ファイルドロップ機能が有効化されました")
            
        except ImportError:
            # tkinterdnd2がインストールされていない場合は何もしない
            pass
        except AttributeError:
            # 必要な属性/メソッドが存在しない場合は何もしない
            pass
        except Exception as e:
            # その他のエラーの場合のみログ出力（デバッグレベル）
            logger.debug(f"ファイルドロップ機能は利用できません: {e}")
    
    def handle_dropped_file(self, file_path):
        """🎯 ドロップされたファイルの処理"""
        try:
            import os
            file_ext = os.path.splitext(file_path.lower())[1]
            
            if file_ext == '.json':
                # JSONファイル -> 設定読み込み
                self.app.load_config_file(file_path)
                self.app.update_status(f"⚡ 設定ファイルを読み込みました: {os.path.basename(file_path)}")
            elif file_ext in ['.png', '.jpg', '.jpeg', '.bmp', '.gif']:
                # 画像ファイル -> 画像クリックステップ追加
                step = Step("image_click", params={
                    "path": file_path,
                    "threshold": 0.8,
                    "click_type": "single", 
                    "retry": 3,
                    "delay": 1.0
                }, comment=f"画像クリック: {os.path.basename(file_path)}")
                self.app.add_step(step)
                self.app.update_status(f"⚡ 画像クリックを追加: {os.path.basename(file_path)}")
            else:
                self.app.update_status(f"⚠️ サポートされていないファイル形式: {file_ext}")
                
        except Exception as e:
            logger.error(f"ファイル処理エラー: {file_path}, エラー: {e}")
            self.app.update_status(f"❌ ファイル処理エラー: {os.path.basename(file_path)}")
    
    def show_context_menu(self, event):
        """コンテキストメニューを表示"""
        context_menu = tk.Menu(self.tree, tearoff=0)
        context_menu.configure(bg=AppConfig.THEME['bg_secondary'],
                             fg=AppConfig.THEME['fg_primary'],
                             activebackground=AppConfig.THEME['bg_hover'],
                             activeforeground=AppConfig.THEME['fg_primary'])
        
        context_menu.add_command(label="編集", command=self.app.edit_selected_step)
        context_menu.add_command(label="コピー", command=self.app.copy_selected_step)
        context_menu.add_command(label="貼り付け", command=self.app.paste_step)
        context_menu.add_separator()
        context_menu.add_command(label="上に移動", command=lambda: self.app.move_up())
        context_menu.add_command(label="下に移動", command=lambda: self.app.move_down())
        context_menu.add_separator()
        context_menu.add_command(label="有効/無効切り替え", command=lambda: self.app.toggle_step_enabled(self.tree.index(self.tree.selection()[0])))
        context_menu.add_command(label="削除", command=self.app.delete_selected)
        
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

class ModernDialog:
    """再設計されたモダンなデザインのカスタムダイアログ"""
    def __init__(self, parent: tk.Tk, title: str, fields: List[Dict[str, Any]], 
                 width: int = 700, height: int = 800):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry(f"{width}x{height}")
        self.dialog.resizable(True, True)
        self.dialog.minsize(500, 400)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.center_window()
        
        self.dialog.configure(bg=AppConfig.THEME['bg_primary'])
        
        self.result = None
        self.fields = fields
        self.entries: Dict[str, ttk.Entry | ttk.Combobox | tk.Text | ttk.Checkbutton] = {}
        self.field_frames: Dict[str, ttk.Frame] = {}
        self.setup_ui()
        
        self.dialog.bind('<Escape>', lambda e: self.dialog.destroy())
        
    def center_window(self):
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f'+{x}+{y}')

    def setup_ui(self):
        """再設計されたモダンなUIを構築"""
        # --- ボタンフレーム（ダイアログ下部に固定） ---
        button_frame = ttk.Frame(self.dialog, style='Toolbar.TFrame')
        button_frame.pack(side='bottom', fill='x', padx=20, pady=(10, 20), anchor='s')
        
        ok_button = ttk.Button(button_frame, text="OK", command=self.submit, style='Primary.TButton', width=10)
        ok_button.pack(side='right', padx=(10, 0))
        
        cancel_button = ttk.Button(button_frame, text="キャンセル", command=self.dialog.destroy, style='Modern.TButton', width=10)
        cancel_button.pack(side='right')

        # --- メインフレーム ---
        main_frame = ttk.Frame(self.dialog, style='Modern.TFrame')
        main_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # --- スクロール可能エリア ---
        canvas = tk.Canvas(main_frame, bg=AppConfig.THEME['bg_primary'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview, style='Modern.Vertical.TScrollbar')
        
        content_frame = ttk.Frame(canvas, style='Modern.TFrame')
        
        content_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas_window = canvas.create_window((0, 0), window=content_frame, anchor="nw")
        
        def configure_canvas_window(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", configure_canvas_window)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # --- コンテンツ ---
        title_label = ttk.Label(content_frame, text=self.dialog.title(), style='Subtitle.TLabel')
        title_label.pack(pady=(20, 25), anchor='center')

        for field in self.fields:
            card = ttk.Frame(content_frame, style='Card.TFrame')
            card.pack(fill='x', padx=20, pady=7)
            self.field_frames[field["key"]] = card
            
            inner_frame = ttk.Frame(card, style='Modern.TFrame')
            inner_frame.pack(fill='x', expand=True, padx=15, pady=12)

            label = ttk.Label(inner_frame, text=field["label"], style='Heading.TLabel')
            label.pack(anchor='w', pady=(0, 8))

            widget_frame = ttk.Frame(inner_frame, style='Modern.TFrame')
            widget_frame.pack(fill='x', expand=True)

            if field.get("type") == "combobox":
                widget = ttk.Combobox(widget_frame, values=field.get("values", []), font=AppConfig.FONTS['body'], style='TCombobox')
                widget.set(field.get("default", ""))
                if field.get("on_change"):
                    widget.bind("<<ComboboxSelected>>", lambda e, f=field: self.handle_field_change(f, e))
            
            elif field.get("type") == "bool":
                var = tk.BooleanVar(value=field.get("default", False))
                widget = ttk.Checkbutton(widget_frame, variable=var, style='Modern.TCheckbutton', text=field.get('text', ''))
                widget.var = var
                widget.widget_type = "bool"

            elif field.get("type") == "text":
                text_frame = ttk.Frame(widget_frame)
                text_frame.pack(fill='both', expand=True, pady=(3, 0))
                
                widget = tk.Text(text_frame, 
                                 font=AppConfig.FONTS['code'], 
                                 height=field.get("height", 8), 
                                 wrap='word',
                                 bg=AppConfig.THEME['bg_secondary'],
                                 fg=AppConfig.THEME['fg_secondary'],
                                 insertbackground=AppConfig.THEME['fg_primary'],
                                 relief='solid',
                                 borderwidth=1,
                                 bd=1,
                                 highlightthickness=1,
                                 highlightcolor=AppConfig.THEME['border'],
                                 highlightbackground=AppConfig.THEME['border'],
                                 selectbackground=AppConfig.THEME['selection'],
                                 selectforeground=AppConfig.THEME['fg_primary'])
                
                default_text = str(field.get("default", "") or field.get("placeholder", ""))
                if default_text:
                    widget.insert('1.0', default_text)
                
                text_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=widget.yview, style='Modern.Vertical.TScrollbar')
                widget.configure(yscrollcommand=text_scrollbar.set)
                
                widget.pack(side='left', fill='both', expand=True)
                text_scrollbar.pack(side='right', fill='y')

            else: # entry
                widget = ttk.Entry(widget_frame, font=AppConfig.FONTS['body'], style='Modern.TEntry')
                default_value = field.get("default", "")
                if isinstance(default_value, bool): default_value = str(default_value)
                widget.insert(0, str(default_value))
                widget.widget_type = "entry"

            if field.get("type") not in ["text", "bool"]:
                widget.pack(fill='x', pady=(3, 0))
            elif field.get("type") == "bool":
                widget.pack(anchor='w', pady=(3, 0))
            self.entries[field["key"]] = widget

            if 'help' in field:
                help_label = ttk.Label(inner_frame, text=field['help'], style='Caption.TLabel', wraplength=550)
                help_label.pack(anchor='w', pady=(8, 0))

        if self.entries:
            first_entry = next(iter(self.entries.values()))
            first_entry.focus_set()
        
        self.update_field_visibility()

    def handle_field_change(self, field, event):
        """フィールド変更時の処理"""
        self.update_field_visibility()
    
    def update_field_visibility(self):
        """フィールドの表示/非表示を更新"""
        for field in self.fields:
            if field.get("show_condition"):
                condition = field["show_condition"]
                show = self.evaluate_condition(condition)
                frame = self.field_frames[field["key"]]
                
                if show:
                    frame.pack(fill='x', padx=20, pady=7)
                else:
                    frame.pack_forget()
    
    def evaluate_condition(self, condition):
        """条件を評価"""
        field_key = condition["field"]
        expected_value = condition["value"]
        
        if field_key in self.entries:
            widget = self.entries[field_key]
            current_value = widget.get() if hasattr(widget, 'get') else ""
            return current_value == expected_value
        return False

    def submit(self):
        """入力値を検証して結果を返す"""
        try:
            result = {}
            for field in self.fields:
                # 非表示のフィールドは送信データに含めない
                if field.get("show_condition"):
                    condition = field["show_condition"]
                    if not self.evaluate_condition(condition):
                        continue

                widget = self.entries[field["key"]]
                if field.get("type") == "text":
                    value = widget.get('1.0', 'end-1c').strip()
                elif field.get("type") == "bool":
                    value = widget.var.get()
                else:
                    value = widget.get().strip()
                
                if field.get("required", False) and not value:
                    raise ValueError(f"「{field['label']}」は必須入力です")
                
                if value:
                    if field.get("type") == "float":
                        value = float(value)
                        if "min" in field and value < field["min"]: raise ValueError(f"「{field['label']}」は{field['min']}以上で入力してください")
                        if "max" in field and value > field["max"]: raise ValueError(f"「{field['label']}」は{field['max']}以下で入力してください")
                    elif field.get("type") == "int":
                        value = int(value)
                        if "min" in field and value < field["min"]: raise ValueError(f"「{field['label']}」は{field['min']}以上で入力してください")
                        if "max" in field and value > field["max"]: raise ValueError(f"「{field['label']}」は{field['max']}以下で入力してください")
                
                result[field["key"]] = value
            
            self.result = result
            self.dialog.destroy()
            
        except ValueError as e:
            self.show_error_with_sound("入力エラー", str(e))
            logger.error(f"入力エラー: {e}")
        except Exception as e:
            self.show_error_with_sound("未知のエラー", f"予期しないエラーが発生しました: {e}")
            logger.error(f"ダイアログエラー: {e}")

    def get_result(self) -> Optional[Dict[str, Any]]:
        """ダイアログの結果を取得"""
        self.dialog.wait_window()
        return self.result
    
    def show_error_with_sound(self, title, message):
        """エラーメッセージを音付きで表示"""
        self.dialog.bell()
        messagebox.showerror(title, message, parent=self.dialog)

class AutoActionTool:
    """メインアプリケーションクラス"""
    
    def __init__(self, root: tk.Tk):
        # 必要なディレクトリを作成
        AppConfig.ensure_directories()
        
        # アプリケーションの初期化
        self.root = root
        self.root.title(f"{AppConfig.APP_NAME} v{AppConfig.VERSION}")
        
        # ダークテーマを適用
        AppConfig.apply_dark_theme(self.root)
        
        # Animation system initialization
        self.animation_queue = []
        self.animation_running = False
        self.hover_animations = {}
        
        # Progress visualization system
        self.execution_stats = {
            'total_steps': 0,
            'completed_steps': 0,
            'start_time': None,
            'estimated_time': None,
            'success_count': 0,
            'error_count': 0,
            'current_step_name': ''
        }
        
        # アイコン設定（利用可能な場合）
        try:
            self.root.iconbitmap(default='icon.ico')
        except:
            pass  # アイコンがない場合は無視
        
        # メンバ変数の初期化
        self.steps: List[Step] = []
        self.running = False
        self.config_file = AppConfig.DEFAULT_CONFIG_FILE
        self.selected_monitor = 0
        self.loop_count = 1
        self.auto_save_enabled = True
        # last_execution_time 変数を削除（未使用）
        self.execution_history = []
        
        # 部分繰り返し実行機能
        # 繰り返し実行は新しいアクション方式を使用
        
        # システム情報の取得
        self.monitors = get_monitors()
        self.dpi_scale = self.get_dpi_scale()
        
        # pyautogui設定
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.1
        
        # クリップボード関連
        self.clipboard_step = None
        
        # 古いUndo/Redo機能は削除（新しいシステムは後で初期化される）
        
        # 実行状態管理
        self.execution_log = []
        self.execution_start_index = 0
        
        # UI変数の初期化
        self.progress_var = tk.StringVar(value="")
        
        # キャプチャ機能
        self.capture_window = None
        self.is_capturing = False
        
        # GUIの初期化（よりコンパクトサイズ）
        self.root.geometry("1080x720")  # 横幅をさらに10%縮小（1200→1080）
        self.root.minsize(900, 650)    # 最小サイズも比例調整
        
        # ウィンドウを中央に配置
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 1080) // 2
        y = (screen_height - 730) // 2
        self.root.geometry(f"1080x730+{x}+{y}")
        self.setup_modern_gui()
        
        # 最終設定を読み込み
        self.load_last_config()
        
        # ウィンドウ終了時の処理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        logger.info(f"アプリケーション初期化完了: モニター数={len(self.monitors)}, DPIスケーリング={self.dpi_scale}%")

    def get_dpi_scale(self) -> float:
        """現在のDPIスケーリングを取得"""
        try:
            if sys.platform == "win32":
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
                hdc = ctypes.windll.user32.GetDC(0)
                dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
                ctypes.windll.user32.ReleaseDC(0, hdc)
                return (dpi / 96.0) * 100
            else:
                # Linux/macOSの場合はデフォルト値を返す
                return AppConfig.DEFAULT_DPI_SCALE
        except (OSError, ctypes.WinError, AttributeError) as e:
            logger.warning(f"DPIスケーリング取得に失敗（詳細エラー）: {e}")
            return AppConfig.DEFAULT_DPI_SCALE
        except Exception as e:
            logger.error(f"DPIスケーリング取得で予期しないエラー: {e}")
            return AppConfig.DEFAULT_DPI_SCALE
    
    def on_closing(self):
        """アプリケーション終了時の処理"""
        try:
            if self.running:
                if messagebox.askyesno("終了確認", "実行中です。アプリケーションを終了しますか？"):
                    self.stop_execution()
                    self.save_last_config()
                    self.root.destroy()
            else:
                self.save_last_config()
                self.root.destroy()
        except Exception as e:
            logger.error(f"アプリケーション終了処理でエラー: {e}")
            # 致命的エラーでもアプリケーションは終了
            try:
                self.root.destroy()
            except:
                sys.exit(1)

    def setup_modern_gui(self):
        """次世代プロフェッショナルUIを構築"""
        
        # メインコンテナ
        main_container = ttk.Frame(self.root, style='Modern.TFrame')
        main_container.pack(fill='both', expand=True, padx=20, pady=20)
        
        # プロフェッショナルステータスバー（先に作成して最下部を確保）
        self.setup_premium_status_bar(main_container)
        
        # ヘッダーエリア（新デザイン）
        self.setup_premium_header(main_container)
        
        # メインレイアウトエリア（pady縮小）
        layout_frame = ttk.Frame(main_container, style='Modern.TFrame')
        layout_frame.pack(fill='both', expand=True, pady=10)  # pady 20→10に縮小
        
        # 3カラムレイアウトを作成
        self.setup_three_column_layout(layout_frame)
        
        
        # 高級ホットキーとシステム初期化
        self.setup_advanced_features()
        
        # レイアウト調整を適用
        self.root.after(200, self.apply_layout_spacing)
        
    def setup_three_column_layout(self, parent):
        """確実な4:3:3比率レイアウトを構築（Grid Layout使用）"""
        # メインコンテナフレーム
        main_container = ttk.Frame(parent, style='Modern.TFrame')
        main_container.pack(fill='both', expand=True, padx=3, pady=3)
        
        # レイアウトマネージャーをpackに統一し、不整合を解消します。
        # 右から配置していくことで、元のレイアウトを再現します。
        right_panel = self.create_preview_control_panel(main_container)
        right_panel.pack(side='right', fill='both', expand=True, padx=(2, 0))

        center_panel = self.create_tool_palette_panel(main_container)
        center_panel.pack(side='right', fill='both', expand=True, padx=1)

        left_panel = self.create_step_management_panel(main_container)
        left_panel.pack(side='right', fill='both', expand=True, padx=(0, 2))
        
        self.content_frame = main_container
        self.main_container = main_container
        
    def setup_advanced_features(self):
        """高級機能とシステムを初期化"""
        # ホットキーを設定
        self.setup_hotkeys()
        
        
        # Undo/Redo システム (最大50段階)
        self.undo_stack = deque(maxlen=50)
        self.redo_stack = deque(maxlen=50)
        self.current_state_id = None
    
    def apply_layout_spacing(self, scale: float = None):
        """レイアウトの余白を調整"""
        if scale is None:
            scale = AppConfig.SPACING_SCALE
        
        try:
            # パディング値を計算
            base_padx = int(4 * scale)
            base_pady = int(4 * scale)
            
            # 主要ブロック間の余白を調整（安全なもののみ）
            if hasattr(self, 'drag_drop_tree') and hasattr(self.drag_drop_tree, 'tree'):
                try:
                    # ツリーはgrid管理のため、paddingはgrid_configureで調整
                    self.drag_drop_tree.tree.grid_configure(pady=(base_pady + 4, base_pady))
                except Exception:
                    pass  # 失敗しても処理を継続
                
        except Exception as e:
            logger.warning(f"レイアウト調整エラー: {e}")
    
    def _adjust_widget_spacing(self, widget, scale: float):
        """ウィジェットの余白を再帰的に調整"""
        try:
            # フレームの余白を調整
            if isinstance(widget, (ttk.Frame, tk.Frame)):
                current_padx = widget.pack_info().get('padx', 0) 
                current_pady = widget.pack_info().get('pady', 0)
                
                if current_padx or current_pady:
                    new_padx = int((current_padx if current_padx else 0) * scale)
                    new_pady = int((current_pady if current_pady else 0) * scale)
                    widget.pack_configure(padx=new_padx, pady=new_pady)
            
            # ボタンのタッチターゲットを拡張
            elif isinstance(widget, (ttk.Button, tk.Button)):
                current_padx = widget.pack_info().get('padx', 0)
                current_pady = widget.pack_info().get('pady', 0)
                
                enhanced_padx = int((current_padx if current_padx else 2) + 2 * scale)
                enhanced_pady = int((current_pady if current_pady else 2) + 2 * scale) 
                widget.pack_configure(padx=enhanced_padx, pady=enhanced_pady)
            
            # 子ウィジェットを再帰的に処理
            for child in widget.winfo_children():
                self._adjust_widget_spacing(child, scale)
                
        except Exception:
            # エラーが発生しても処理を継続
            pass
    
    def setup_premium_header(self, parent):
        """プレミアムヘッダーデザインを構築"""
        # ヘッダーカード
        header_card = ttk.Frame(parent, style='Card.TFrame')
        header_card.pack(fill='x', pady=(0, 15))
        
        # ヘッダー内部コンテンツ
        header_content = ttk.Frame(header_card, style='Modern.TFrame')
        header_content.pack(fill='x', padx=20, pady=15)
        
        # 左側: アプリタイトルとバージョン
        left_header = ttk.Frame(header_content, style='Modern.TFrame')
        left_header.pack(side='left', fill='both', expand=True)
        
        # メインタイトル
        title_label = ttk.Label(left_header, 
                               text=f"🚀 {AppConfig.APP_NAME}",
                               style='Title.TLabel')
        title_label.pack(anchor='w')
                
        # バージョンとステータス
        info_frame = ttk.Frame(left_header, style='Modern.TFrame')
        info_frame.pack(anchor='w', fill='x', pady=(5, 0))
        
        version_label = ttk.Label(info_frame, 
                                 text=f"v{AppConfig.VERSION}",
                                 style='Subtitle.TLabel')
        version_label.pack(side='left')
        
        # セパレータ
        ttk.Label(info_frame, text="  |  ", 
                 style='Caption.TLabel').pack(side='left')
        
        # システム情報
        self.system_info_label = ttk.Label(info_frame, 
                                          text="💻 システム準備完了", 
                                          style='Caption.TLabel')
        self.system_info_label.pack(side='left')
        
        # 右側: クイックアクションボタン
        right_header = ttk.Frame(header_content, style='Modern.TFrame')
        right_header.pack(side='right')
        
        
        # ボタンの参照を保存するために個別に作成
        self.run_button = ttk.Button(right_header, text="⚡ 実行", command=self.run_all_steps, style='Primary.TButton')
        self.run_button.pack(side='left', padx=(0, 10))
        self.setup_hover_animations(self.run_button)
        
        self.stop_button = ttk.Button(right_header, text="⏸️ 停止 (ESC)", command=self.stop_execution, style='Danger.TButton', state='disabled')
        self.stop_button.pack(side='left', padx=(0, 10))
        self.setup_hover_animations(self.stop_button)
    
    def update_execution_buttons(self, running):
        """実行状態に応じてボタンの有効/無効を切り替え"""
        try:
            if hasattr(self, 'run_button'):
                if running:
                    self.run_button.configure(state='disabled')
                    self.stop_button.configure(state='normal')
                else:
                    self.run_button.configure(state='normal')
                    self.stop_button.configure(state='disabled')
        except Exception as e:
            logger.debug(f"ボタン状態更新エラー: {e}")
            
    def create_step_management_panel(self, parent):
        """ステップ管理パネルを作成"""
        # メインパネル
        step_panel = ttk.Frame(parent, style='Card.TFrame')
        
        # パネルヘッダー
        header = ttk.Frame(step_panel, style='Toolbar.TFrame')
        header.pack(fill='x', padx=10, pady=(10, 0))
        
        # タイトルとアイコン
        title_frame = ttk.Frame(header, style='Toolbar.TFrame')
        title_frame.pack(side='left', fill='x', expand=True)
        
        ttk.Label(title_frame, text="📋 ステップリスト管理", 
                 style='Heading.TLabel').pack(side='left', anchor='w')
        
        # ステップ統計ラベルを削除（未更新のため不要）
        
        # 右側ツールバー
        toolbar = ttk.Frame(header, style='Toolbar.TFrame')
        toolbar.pack(side='right')
        
        # 主要編集ボタン（文字表示）
        edit_buttons = [
            ("編集", self.edit_selected_step, "選択したステップを編集", 4),
            ("削除", self.delete_selected, "選択したステップを削除", 4), 
            ("コピー", self.copy_selected_step, "選択したステップをコピー", 4),
            ("貼付", self.paste_step, "コピーしたステップを貼り付け", 4),
            ("クリア", self.clear_all_steps, "すべてのステップをクリア", 4)  # "全削除"→"クリア"に短縮
        ]
        
        for text, command, tooltip, width in edit_buttons:
            btn = ttk.Button(toolbar, text=text, command=command, 
                           style='Compact.TButton', width=width)
            btn.pack(side='left', padx=1, pady=1, ipady=0)  # ipady=0で縦を縮小
            self.add_tooltip(btn, tooltip)
        
        # ステップリストコンテンツエリア（コンパクト）
        content_frame = ttk.Frame(step_panel, style='Modern.TFrame')
        content_frame.pack(fill='both', expand=True, padx=8, pady=8)
        
        # ドラッグ&ドロップ対応ツリービュー
        self.drag_drop_tree = DragDropTreeview(content_frame, self)
        self.tree = self.drag_drop_tree.tree
        
        # スクロールバー（モダンスタイル）
        v_scroll = ttk.Scrollbar(content_frame, orient="vertical", 
                                command=self.tree.yview, style='Modern.Vertical.TScrollbar')
        h_scroll = ttk.Scrollbar(content_frame, orient="horizontal", 
                                command=self.tree.xview, style='Modern.Horizontal.TScrollbar')
        
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        # グリッド配置
        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(0, weight=1)
        
        
        return step_panel
        
    def create_tool_palette_panel(self, parent):
        """ツールパレット+繰り返し設定パネルを作成"""
        # メインパネル
        tool_panel = ttk.Frame(parent, style='Card.TFrame')
        
        # ツールパレットセクション
        tools_header = ttk.Frame(tool_panel, style='Toolbar.TFrame')
        tools_header.pack(fill='x', padx=8, pady=(8, 0))
        
        ttk.Label(tools_header, text="🛠️ツール", 
                 style='Heading.TLabel').pack(anchor='w')
        
        # ツールボタンコンテンツ（縦並び）
        tools_content = ttk.Frame(tool_panel, style='Modern.TFrame')
        tools_content.pack(fill='both', expand=True, padx=8, pady=(5, 10))
        
        # ツールボタン縦並び配置
        self.setup_vertical_tools(tools_content)
        
        return tool_panel
        
    def setup_compact_tools(self, parent):
        """コンパクトツールボタンを作成"""
        # 主要ツールのみコンパクト表示
        compact_tools = [
            ("画像", self.add_step_image_click, "画像をクリックするステップを追加"),
            ("座標", self.add_step_coord_click, "指定座標をクリックするステップを追加"), 
            ("オフセ", self.add_step_image_relative_right_click, "画像基準でオフセット右クリック"),
            ("待機", self.add_step_sleep, "指定時間待機するステップを追加"),
            ("キー", self.add_step_key_custom, "キーボード入力ステップを追加"),
            ("文字", self.add_step_custom_text, "文字列入力ステップを追加"),
            ("開始", self.add_step_repeat_start, "繰り返し処理の開始点"),
            ("終了", self.add_step_repeat_end, "繰り返し処理の終了点"),
            ("ドラッグ", self.add_step_coord_drag, "座標間ドラッグするステップを追加")
        ]
        
        # 2行4列のグリッド配置
        for i, (icon, command, tooltip) in enumerate(compact_tools):
            row = i // 4
            col = i % 4
            btn = ttk.Button(parent, text=icon, command=command,
                           style='Modern.TButton', width=4)
            btn.grid(row=row, column=col, padx=2, pady=2, sticky='ew')
            self.add_tooltip(btn, tooltip)
        
        # 列の重み設定
        for i in range(4):
            parent.columnconfigure(i, weight=1)
    
    def setup_vertical_tools(self, parent):
        """役割ごとにグループ化されたカード方式ツールボタンを配置"""
        # 役割ごとにグループ化されたツールリスト
        tool_groups = [
            # クリック系操作
            {
                "title": "🖱️ クリック操作",
                "tools": [
                    ("🖱️", "画像クリック", self.add_step_image_click, "画像をクリックするステップを追加", "Click.TButton"),
                    ("📍", "座標クリック", self.add_step_coord_click, "指定座標をクリックするステップを追加", "Click.TButton"),
                    ("↗️", "オフセットクリック", self.add_step_image_relative_right_click, "画像からオフセットして右クリック", "Click.TButton"),
                    ("↔️", "座標ドラッグ", self.add_step_coord_drag, "座標間ドラッグするステップを追加", "Click.TButton")
                ]
            },
            # 待機系操作
            {
                "title": "⏰ 待機操作",
                "tools": [
                    ("⏱️", "待機(スリープ)", self.add_step_sleep_seconds, "指定秒数間待機するステップを追加", "Wait.TButton"),
                    ("🕐", "待機(時間指定)", self.add_step_sleep_time, "指定時刻まで待機するステップを追加", "Wait.TButton")
                ]
            },
            # 入力系操作
            {
                "title": "⌨️ 入力操作",
                "tools": [
                    ("⌨️", "キー入力", self.add_step_key_custom, "キーボード入力ステップを追加", "Input.TButton"),
                    ("📝", "文字入力", self.add_step_custom_text, "文字列入力ステップを追加", "Input.TButton"),
                    ("⚙️", "コマンド実行", self.add_step_cmd_command, "コマンド実行ステップを追加", "System.TButton")
                ]
            },
            # 制御系操作
            {
                "title": "🔄 制御操作",
                "tools": [
                    ("🔄", "繰り返し開始", self.add_step_repeat_start, "繰り返し処理の開始点を設定", "Repeat.TButton"),
                    ("🔚", "繰り返し終了", self.add_step_repeat_end, "繰り返し処理の終了点を設定", "Repeat.TButton")
                ]
            }
        ]
        
        current_row = 0
        
        for group in tool_groups:
            # グループヘッダー
            if current_row > 0:
                # セパレータ（空白行）
                current_row += 1
            
            header_frame = ttk.Frame(parent, style='Toolbar.TFrame')
            header_frame.grid(row=current_row, column=0, columnspan=2, sticky='ew', pady=(5, 2))
            
            ttk.Label(header_frame, text=group["title"], 
                     style='Heading.TLabel', font=AppConfig.FONTS['small']).pack(anchor='w')
            current_row += 1
            
            # グループ内のツールを2列配置
            tools = group["tools"]
            group_start_row = current_row
            
            for i, (icon, text, command, tooltip, style) in enumerate(tools):
                row = group_start_row + (i // 2)
                col = i % 2
                
                # カードフレーム
                card_frame = ttk.Frame(parent, style='HighlightCard.TFrame')
                card_frame.grid(row=row, column=col, padx=2, pady=2, sticky='ew')
                
                # カードボタン
                btn = ttk.Button(card_frame, text=f"{icon}\n{text}", command=command,
                               style=style, width=12)
                btn.pack(fill='x', padx=3, pady=3)
                
                # ホバーアニメーション追加
                self.setup_hover_animations(btn)
                self.add_tooltip(btn, tooltip)
            
            # 次のグループ用に行を更新
            current_row = group_start_row + ((len(tools) - 1) // 2) + 1
        
        # 列の重み設定（均等配置）
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)
    

    def create_preview_control_panel(self, parent):
        """プレビュー・コントロールパネルを作成"""
        # メインパネル
        preview_panel = ttk.Frame(parent, style='Card.TFrame')
        
        # 上部: 画像プレビューエリア
        preview_header = ttk.Frame(preview_panel, style='Toolbar.TFrame')
        preview_header.pack(fill='x', padx=10, pady=(10, 0))
        
        ttk.Label(preview_header, text="🖼️プレビュー", 
                 style='Heading.TLabel').pack(anchor='w', pady=(0, 4))  # 下部余白追加
        
        # 画像プレビューウィジェット
        self.image_preview = ImagePreviewWidget(preview_panel, self)
        self.image_preview.frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 中部: コントロールエリア
        control_header = ttk.Frame(preview_panel, style='Toolbar.TFrame')
        control_header.pack(fill='x', padx=10, pady=(10, 0))
        
        ttk.Label(control_header, text="🎮 コントロール", 
                 style='Heading.TLabel').pack(anchor='w', pady=(0, 4))  # 下部余白追加
        
        # コントロールパネル
        self.setup_enhanced_control_panel(preview_panel)
        
        # プログレス機能はステータスバーに統合済み
        
        return preview_panel
    
    def setup_categorized_tools(self, parent):
        """カテゴリ別ツールボタンを設定"""
        # スクロール可能なフレーム
        canvas = tk.Canvas(parent, bg=AppConfig.THEME['bg_secondary'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview,
                                 style='Modern.Vertical.TScrollbar')
        scrollable_frame = ttk.Frame(canvas, style='Modern.TFrame')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # カテゴリ別ツールグループ（統一アイコン使用）
        categories = [
            ("🎯 クリック操作", [
                (f"{AppConfig.get_step_icon('image_click')} 画像クリック", self.add_step_image_click, AppConfig.CATEGORY_COLORS['click']),
                (f"{AppConfig.get_step_icon('coord_click')} 座標クリック", self.add_step_coord_click, AppConfig.CATEGORY_COLORS['click']),
                (f"{AppConfig.get_step_icon('image_relative_right_click')} 画像オフセット", self.add_step_image_relative_right_click, AppConfig.CATEGORY_COLORS['click']),
                (f"{AppConfig.get_step_icon('coord_drag')} 座標ドラッグ", self.add_step_coord_drag, AppConfig.CATEGORY_COLORS['click'])
            ]),
            ("⌨️ 入力操作", [
                (f"{AppConfig.get_step_icon('key')} キー入力", self.add_step_key_custom, AppConfig.CATEGORY_COLORS['input']),
                (f"{AppConfig.get_step_icon('custom_text')} テキスト", self.add_step_custom_text, AppConfig.CATEGORY_COLORS['input'])
            ]),
            ("⏱️ 制御操作", [
                (f"{AppConfig.get_step_icon('sleep')} 待機", self.add_step_sleep, AppConfig.CATEGORY_COLORS['wait']),
                (f"{AppConfig.get_step_icon('cmd_command')} コマンド", self.add_step_cmd_command, AppConfig.CATEGORY_COLORS['system'])
            ]),
            ("🔁 繰り返し", [
                (f"{AppConfig.get_step_icon('repeat_start')} 繰り返し開始", self.add_step_repeat_start, AppConfig.CATEGORY_COLORS['repeat']),
                (f"{AppConfig.get_step_icon('repeat_end')} 繰り返し終了", self.add_step_repeat_end, AppConfig.CATEGORY_COLORS['repeat'])
            ])
        ]
        
        for i, (category_title, tools) in enumerate(categories):
            # セパレータを追加（最初のカテゴリ以外）
            if i > 0:
                separator = ttk.Separator(scrollable_frame, orient='horizontal')
                separator.pack(fill='x', padx=10, pady=(6, 6))
            
            # カテゴリヘッダー
            category_frame = ttk.LabelFrame(scrollable_frame, text=category_title,
                                          style='Modern.TFrame')
            category_frame.pack(fill='x', padx=5, pady=(6, 8))  # 上部余白を追加
            
            # ツールボタン
            for tool_name, command, color in tools:
                btn_frame = ttk.Frame(category_frame, style='Modern.TFrame')
                btn_frame.pack(fill='x', padx=5, pady=2)
                
                btn = ttk.Button(btn_frame, text=tool_name, command=command,
                               style='Modern.TButton')
                btn.pack(fill='x')
        
        # キャンバスとスクロールバーを配置
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    # === Animation System ===
    def animate_widget_fade(self, widget, start_alpha=0.0, end_alpha=1.0, duration=300, callback=None):
        """ウィジェットのフェードアニメーション"""
        try:
            steps = 20
            step_duration = duration // steps
            alpha_step = (end_alpha - start_alpha) / steps
            
            def fade_step(current_step=0):
                if current_step <= steps:
                    alpha = start_alpha + (alpha_step * current_step)
                    # tkinterではalpha直接制御は限定的なので、色で代用
                    try:
                        if hasattr(widget, 'configure'):
                            # ボタンの場合
                            if 'Button' in str(type(widget)):
                                current_color = AppConfig.THEME['bg_primary']
                                widget.configure(relief='raised' if alpha > 0.5 else 'flat')
                    except:
                        pass
                    
                    self.root.after(step_duration, lambda: fade_step(current_step + 1))
                elif callback:
                    callback()
            
            fade_step()
        except Exception as e:
            logger.error(f"フェードアニメーションエラー: {e}")

    def animate_widget_slide(self, widget, start_x, end_x, duration=300, callback=None):
        """ウィジェットのスライドアニメーション"""
        try:
            steps = 15
            step_duration = duration // steps
            x_step = (end_x - start_x) / steps
            
            def slide_step(current_step=0):
                if current_step <= steps:
                    x = start_x + (x_step * current_step)
                    widget.place_configure(x=int(x))
                    self.root.after(step_duration, lambda: slide_step(current_step + 1))
                elif callback:
                    callback()
            
            slide_step()
        except Exception as e:
            logger.error(f"スライドアニメーションエラー: {e}")

    def animate_button_pulse(self, button, color_start=None, color_end=None, duration=200):
        """ボタンのパルスアニメーション"""
        try:
            if not color_start:
                color_start = AppConfig.THEME['bg_primary']
            if not color_end:
                color_end = AppConfig.THEME['bg_accent']
                
            original_relief = button.cget('relief')
            
            # パルス効果
            button.configure(background=color_end, relief='raised')
            self.root.after(duration // 2, lambda: button.configure(
                background=color_start, relief=original_relief))
        except Exception as e:
            logger.error(f"パルスアニメーションエラー: {e}")

    def setup_hover_animations(self, widget, hover_color=None, normal_color=None):
        """ホバーアニメーションをセットアップ"""
        try:
            if not hover_color:
                hover_color = AppConfig.THEME['bg_hover']
            if not normal_color:
                normal_color = AppConfig.THEME['bg_primary']
                
            def on_enter(event):
                try:
                    widget.configure(background=hover_color)
                    self.animate_button_pulse(widget, normal_color, hover_color, 150)
                except:
                    pass
                    
            def on_leave(event):
                try:
                    widget.configure(background=normal_color)
                except:
                    pass
            
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)
        except Exception as e:
            logger.error(f"ホバーアニメーション設定エラー: {e}")


    def animate_action_feedback(self, action_name):
        """アクション実行時のフィードバックアニメーション"""
        try:
            # ステータスバーに短時間のアクセントカラー表示
            if hasattr(self, 'status_label'):
                original_bg = self.status_label.cget('background')
                original_fg = self.status_label.cget('foreground')
                
                # アクセントカラーで強調
                self.status_label.configure(
                    background=AppConfig.THEME['bg_accent'],
                    foreground=AppConfig.THEME['fg_primary']
                )
                
                # 元に戻す
                self.root.after(800, lambda: self.status_label.configure(
                    background=original_bg, foreground=original_fg))
                    
            # 成功音効果の代わりにビジュアル効果
            logger.info(f"アクション実行: {action_name} - フィードバック完了")
            
        except Exception as e:
            logger.debug(f"アクションフィードバックアニメーションエラー: {e}")

    def animate_tree_update(self):
        """ツリー更新時のスムーズアニメーション"""
        try:
            if hasattr(self, 'tree'):
                # ツリー更新の視覚的フィードバック
                tree = self.tree
                original_bg = tree.cget('selectbackground')
                
                # 短時間のハイライト
                tree.configure(selectbackground=AppConfig.THEME['bg_accent'])
                self.root.after(200, lambda: tree.configure(selectbackground=original_bg))
        except Exception as e:
            logger.debug(f"ツリー更新アニメーションエラー: {e}")

    # === Progress Visualization System ===
    def create_animated_progress_bar(self, parent, width=300, height=20):
        """アニメーション付きプログレスバーを作成"""
        try:
            progress_frame = ttk.Frame(parent, style='Card.TFrame')
            
            # プログレスバーキャンバス
            canvas = tk.Canvas(progress_frame, width=width, height=height, 
                             bg=AppConfig.THEME['bg_secondary'], highlightthickness=1,
                             highlightcolor=AppConfig.THEME['border'])
            canvas.pack(pady=5)
            
            # プログレス情報
            info_label = ttk.Label(progress_frame, text="待機中...", 
                                 style='Modern.TLabel', font=AppConfig.FONTS['small'])
            info_label.pack()
            
            # プログレスバーのデータ
            progress_data = {
                'canvas': canvas,
                'info_label': info_label,
                'width': width,
                'height': height,
                'progress': 0.0,
                'animated': True,
                'gradient_offset': 0
            }
            
            return progress_frame, progress_data
            
        except Exception as e:
            logger.error(f"プログレスバー作成エラー: {e}")
            return None, None

    def update_progress_bar(self, progress_data, progress, text="", animate=True):
        """プログレスバーを更新"""
        try:
            if not progress_data:
                return
                
            canvas = progress_data['canvas']
            width = progress_data['width']
            height = progress_data['height']
            
            # キャンバスをクリア
            canvas.delete("all")
            
            # 背景
            canvas.create_rectangle(2, 2, width-2, height-2, 
                                  fill=AppConfig.THEME['bg_tertiary'], 
                                  outline=AppConfig.THEME['border'])
            
            if progress > 0:
                # プログレスバー幅
                bar_width = (width - 6) * min(progress, 1.0)
                
                if animate and bar_width > 10:
                    # グラデーション効果付きプログレスバー
                    self.draw_gradient_progress(canvas, 3, 3, bar_width, height-6)
                else:
                    # シンプルなプログレスバー
                    canvas.create_rectangle(3, 3, 3 + bar_width, height-3,
                                          fill=AppConfig.THEME['bg_accent'],
                                          outline="")
            
            # プログレス情報を更新
            if text:
                progress_data['info_label'].configure(text=text)
                
            progress_data['progress'] = progress
            
        except Exception as e:
            logger.debug(f"プログレスバー更新エラー: {e}")

    def draw_gradient_progress(self, canvas, x, y, width, height):
        """グラデーション効果のプログレスバーを描画"""
        try:
            # グラデーション効果用の複数の矩形
            gradient_steps = min(int(width / 2), 20)  # 最大20ステップ
            if gradient_steps < 1:
                gradient_steps = 1
                
            step_width = width / gradient_steps
            
            for i in range(gradient_steps):
                # 色の計算（アクセントカラーから少し明るい色へ）
                progress_ratio = i / max(gradient_steps - 1, 1)
                color = self.blend_colors(AppConfig.THEME['bg_accent'], 
                                        AppConfig.THEME['bg_hover'], progress_ratio * 0.3)
                
                step_x = x + (i * step_width)
                canvas.create_rectangle(step_x, y, step_x + step_width + 1, y + height,
                                      fill=color, outline="")
        except Exception as e:
            logger.debug(f"グラデーション描画エラー: {e}")

    def blend_colors(self, color1, color2, ratio):
        """2つの色をブレンド"""
        try:
            # 簡単な色ブレンド（16進数カラーの場合）
            if color1.startswith('#') and color2.startswith('#'):
                r1, g1, b1 = int(color1[1:3], 16), int(color1[3:5], 16), int(color1[5:7], 16)
                r2, g2, b2 = int(color2[1:3], 16), int(color2[3:5], 16), int(color2[5:7], 16)
                
                r = int(r1 + (r2 - r1) * ratio)
                g = int(g1 + (g2 - g1) * ratio)
                b = int(b1 + (b2 - b1) * ratio)
                
                return f"#{r:02x}{g:02x}{b:02x}"
            return color1
        except:
            return color1


    def update_execution_stats(self, step_name="", success=None, error=None):
        """実行統計を更新"""
        try:
            from datetime import datetime
            
            # 現在のステップ名を更新
            if step_name:
                self.execution_stats['current_step_name'] = step_name
            
            # 成功・エラーカウントを更新
            if success is True:
                self.execution_stats['success_count'] += 1
                self.execution_stats['completed_steps'] += 1
            elif error is True:
                self.execution_stats['error_count'] += 1
                self.execution_stats['completed_steps'] += 1
                
            # 進捗情報を更新（経過時間と進捗率）
            if self.execution_stats['start_time']:
                progress = self.execution_stats['completed_steps'] / max(self.execution_stats['total_steps'], 1)
                self.update_realtime_info(step_name, progress)
                
        except Exception as e:
            logger.debug(f"統計更新エラー: {e}")

    def start_execution_tracking(self, total_steps=None):
        """実行追跡を開始"""
        try:
            from datetime import datetime
            
            # 繰り返しを考慮した総実行ステップ数を計算
            if total_steps is None:
                total_steps = self.calculate_total_execution_steps()
            
            self.execution_stats.update({
                'total_steps': total_steps,
                'completed_steps': 0,
                'start_time': datetime.now(),
                'success_count': 0,
                'error_count': 0,
                'current_step_name': ''
            })
            
            self.update_execution_stats()
            logger.info(f"実行追跡開始: {total_steps}ステップ（繰り返し考慮済み）")
            
        except Exception as e:
            logger.error(f"実行追跡開始エラー: {e}")

    def calculate_total_execution_steps(self):
        """繰り返しを考慮した総実行ステップ数を計算（repeat_start/repeat_endも含む）"""
        try:
            # 実行計画を生成してその長さを取得（最も正確）
            execution_plan = self._generate_execution_plan()
            
            # 有効なステップのみをカウント
            valid_execution_count = 0
            for step_index, repeat_iter in execution_plan:
                if step_index < len(self.steps) and self.steps[step_index].enabled:
                    valid_execution_count += 1
                    
            return max(valid_execution_count, 1)  # 最低1ステップは保証
            
        except Exception as e:
            logger.error(f"総実行ステップ数計算エラー: {e}")
            # フォールバック: 有効なステップ数をカウント
            return len([s for s in self.steps if s.enabled])

    def animate_step_completion(self, step_index, success=True):
        """ステップ完了時のアニメーション"""
        try:
            # 既存のハイライト処理
            self.animate_step_highlight(step_index)
            
            # 完了エフェクト
            if success:
                # 成功時の緑色パルス
                color = "#4CAF50"  # 緑色
                self.create_completion_effect(step_index, color, "✅")
            else:
                # エラー時の赤色パルス  
                color = "#F44336"  # 赤色
                self.create_completion_effect(step_index, color, "❌")
                
        except Exception as e:
            logger.debug(f"ステップ完了アニメーションエラー: {e}")

    def create_completion_effect(self, step_index, color, icon):
        """完了エフェクトを作成"""
        try:
            if hasattr(self, 'tree'):
                children = self.tree.get_children()
                if 0 <= step_index < len(children):
                    item_id = children[step_index]
                    
                    # 一時的にアイコンを表示
                    original_values = self.tree.item(item_id, 'values')
                    if original_values:
                        new_values = list(original_values)
                        new_values[0] = icon  # ステータス欄にアイコン
                        self.tree.item(item_id, values=new_values)
                        
                        # 短時間後に元に戻す
                        self.root.after(1500, lambda: self.tree.item(item_id, values=original_values))
        except Exception as e:
            logger.debug(f"完了エフェクト作成エラー: {e}")


    def start_realtime_timer(self):
        """リアルタイム更新タイマーを開始"""
        try:
            self.realtime_timer_active = True
            self.update_realtime_display()
        except Exception as e:
            logger.debug(f"リアルタイムタイマー開始エラー: {e}")

    def update_realtime_display(self):
        """リアルタイム表示を定期的に更新"""
        try:
            if not hasattr(self, 'realtime_timer_active') or not self.realtime_timer_active:
                return
                
            # 経過時間をリアルタイムで更新
            if hasattr(self, 'execution_stats') and self.execution_stats.get('start_time'):
                from datetime import datetime
                elapsed = datetime.now() - self.execution_stats['start_time']
                elapsed_str = f"{elapsed.seconds//60:02d}:{elapsed.seconds%60:02d}"
                
                if hasattr(self, 'realtime_labels') and 'elapsed_time' in self.realtime_labels:
                    self.realtime_labels['elapsed_time'].configure(text=elapsed_str)
            
            # 1秒後に再実行
            if hasattr(self, 'root') and self.root:
                self.root.after(1000, self.update_realtime_display)
                
        except Exception as e:
            logger.debug(f"リアルタイム表示更新エラー: {e}")

    def stop_realtime_timer(self):
        """リアルタイム更新タイマーを停止"""
        try:
            self.realtime_timer_active = False
        except Exception as e:
            logger.debug(f"リアルタイムタイマー停止エラー: {e}")

    def update_realtime_info(self, step_name="", progress=0.0):
        """進捗情報を更新"""
        try:
            if not hasattr(self, 'realtime_labels'):
                return
                
            # 現在のステップ名を更新
            if step_name and 'current_step' in self.realtime_labels:
                # ステップ名を短縮表示
                display_name = step_name[:30] + "..." if len(step_name) > 30 else step_name
                self.realtime_labels['current_step'].configure(text=display_name)
                
            # 経過時間を更新
            if self.execution_stats['start_time'] and 'elapsed_time' in self.realtime_labels:
                from datetime import datetime
                elapsed = datetime.now() - self.execution_stats['start_time']
                elapsed_str = f"{elapsed.seconds//60:02d}:{elapsed.seconds%60:02d}"
                self.realtime_labels['elapsed_time'].configure(text=elapsed_str)
                
            # 進捗率をステータスバーのテキストラベルに更新
            if hasattr(self, 'progress_text_label'):
                if progress >= 1.0:
                    progress_percent = "100%"
                else:
                    progress_percent = f"{progress * 100:.0f}%"
                self.progress_text_label.configure(text=progress_percent)
                
                # ステータスバーの背景色を進捗率に合わせて更新
                if hasattr(self, 'progress_bg_frame') and hasattr(self, 'main_status_label'):
                    # 進捗率に応じて背景の幅を変更（右側のコントロールを避けるため90%まで）
                    limited_progress = min(progress * 0.9, 0.9)
                    self.progress_bg_frame.place_configure(relwidth=limited_progress)
                    
                    # ステータステキストの更新
                    if step_name:
                        self.main_status_label.configure(text=f"⚡ {step_name}")
                    elif progress >= 1.0:
                        self.main_status_label.configure(text="✅ 実行完了")
                    elif progress > 0:
                        self.main_status_label.configure(text="⏳ 実行中...")
                    else:
                        self.main_status_label.configure(text="🟢 システム準備完了")
                    
            # プログレスバーを更新
            if hasattr(self, 'main_progress_bar') and self.main_progress_bar:
                progress_text = f"実行中: {step_name}" if step_name else f"進捗: {progress_percent}"
                self.update_progress_bar(self.main_progress_bar, progress, progress_text, animate=True)
                
        except Exception as e:
            logger.debug(f"進捗情報更新エラー: {e}")

    # setup_integrated_progress_panel関数を削除（ステータスバーに統合済み）

    def setup_enhanced_control_panel(self, parent):
        """拡張コントロールパネルを設定"""
        control_frame = ttk.Frame(parent, style='Modern.TFrame')
        control_frame.pack(fill='x', padx=10, pady=10)
        
        
        # モニターセクション
        monitor_section = ttk.LabelFrame(control_frame, text="モニター",
                                       style='Modern.TLabelframe')
        monitor_section.pack(fill='x', pady=5)
        
        monitor_frame = ttk.Frame(monitor_section, style='Modern.TFrame')
        monitor_frame.pack(fill='x', padx=10, pady=5)
        
        self.monitor_var = tk.StringVar(value="0")
        monitor_combo = ttk.Combobox(monitor_frame, textvariable=self.monitor_var,
                                   values=[f"モニター {i}" for i in range(len(self.monitors))],
                                   width=15, state="readonly", font=AppConfig.FONTS['small'])
        monitor_combo.pack(fill='x', expand=True)
        monitor_combo.bind('<<ComboboxSelected>>', self.on_monitor_selected)
        
        # ファイル操作セクション
        file_section = ttk.LabelFrame(control_frame, text="ファイル",
                                    style='Modern.TLabelframe')
        file_section.pack(fill='x', pady=5)
        
        file_buttons = [
            ("保存", self.save_config),
            ("読込", self.load_config)
        ]
        
        file_grid = ttk.Frame(file_section, style='Modern.TFrame')
        file_grid.pack(fill='x', padx=10, pady=5)
        
        for i, (text, command) in enumerate(file_buttons):
            btn = ttk.Button(file_grid, text=text, command=command,
                           style='Modern.TButton', width=10)
            btn.grid(row=0, column=i, padx=2, sticky='ew')
            
        file_grid.columnconfigure(0, weight=1)
        file_grid.columnconfigure(1, weight=1)
        file_grid.columnconfigure(2, weight=1)
        
        # 設定ファイル選択
        config_select_frame = ttk.Frame(file_section, style='Modern.TFrame')
        config_select_frame.pack(fill='x', padx=10, pady=(5, 0))
        
        ttk.Label(config_select_frame, text="設定", 
                 style='Modern.TLabel', font=AppConfig.FONTS['small']).pack(anchor='w')
        self.config_combo = ttk.Combobox(config_select_frame, state="readonly", width=25,
                                        font=AppConfig.FONTS['small'])
        self.config_combo.pack(fill='x', pady=(2, 0))
        self.config_combo.bind('<<ComboboxSelected>>', self.on_config_selected)
        
        # 初期化後にconfig一覧を更新（遅延実行）
        self.root.after(500, self.update_config_list)
        
        
        
    def setup_premium_status_bar(self, parent):
        """プレミアムステータスバーを設定"""
        # 統合ステータスバー（プログレスバー機能付き）
        self.status_frame = tk.Frame(parent, bg=AppConfig.THEME['bg_secondary'], height=40)
        self.status_frame.pack(fill='x', pady=(10, 0), side='bottom')  # side='bottom'で最下部に固定
        self.status_frame.pack_propagate(False)  # 高さを確実に保持
        
        # コンテンツエリア（文字表示エリア）
        self.status_content = tk.Frame(self.status_frame, bg=AppConfig.THEME['bg_secondary'])
        self.status_content.pack(fill='both', expand=True, padx=15, pady=8)
        
        # メインステータスラベルを先に作成
        self.main_status_label = tk.Label(self.status_content, text="🟢 システム準備完了",
                                         bg=AppConfig.THEME['bg_secondary'],
                                         fg=AppConfig.THEME['fg_primary'],
                                         font=AppConfig.FONTS['small'])
        self.main_status_label.pack(side='left', padx=(5, 0))
        
        # 右側のコントロールエリア
        right_controls = tk.Frame(self.status_content, bg=AppConfig.THEME['bg_secondary'])
        right_controls.pack(side='right', padx=(0, 5))
        
        # ヘルプボタン（?ボタン）
        help_button = tk.Button(right_controls, text="?", 
                               command=self.show_help,
                               bg="#555555",
                               fg=AppConfig.THEME['fg_primary'],
                               font=AppConfig.FONTS['button'],
                               width=3,
                               height=1,
                               relief='flat',
                               cursor='hand2')
        help_button.pack(side='right', padx=(5, 0))
        
        # プログレスパーセンテージラベル
        self.status_progress_label = tk.Label(right_controls, text="0%",
                                            bg=AppConfig.THEME['bg_secondary'],
                                            fg=AppConfig.THEME['fg_primary'],
                                            font=AppConfig.FONTS['button'])
        self.status_progress_label.pack(side='right', padx=(0, 5))
        
        # プログレス背景をラベルの後に作成して最前面に配置
        self.progress_bg_frame = tk.Frame(self.status_content, bg=AppConfig.THEME['bg_accent'])
        self.progress_bg_frame.place(x=0, y=0, relwidth=0, relheight=1)  # 初期は幅0
        self.progress_bg_frame.lower()  # 背景に回す
        
        # realtime_labelsを空の辞書で初期化（他のコードとの互換性のため）
        self.realtime_labels = {}
        
        # プログレス更新用のラベルをステータスバーのものと統一
        self.progress_text_label = self.status_progress_label
    
    # ヘルプとモニター選択のダミー実装
    def show_help_simple(self):
        """簡単なマニュアルを表示"""
        help_text = f"""
{AppConfig.APP_NAME} v{AppConfig.VERSION}
🚀 初心者向けクイックスタートガイド

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ このツールで何ができるの？
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
毎日同じ操作を繰り返していませんか？このツールなら：
• Webサイトのログイン → データ入力 → 保存を自動化
• 定期的なファイルのバックアップ作業を自動化
• 繰り返しのテスト作業を自動化
など、面倒な作業をコンピュータに任せられます！

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ まず試してみよう！（5分チュートリアル）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🔰 Step1: メモ帳を自動で開いてみよう
1. 画面中央の「📷画像クリック」ボタンをクリック
2. 「画像を選択」ボタンを押して、デスクトップのメモ帳アイコンを撮影
3. 「OK」で設定完了 → ステップリスト（左側）に追加されます

🔰 Step2: 文字を自動入力してみよう  
1. 「📝文字入力」ボタンをクリック
2. 「こんにちは！自動化テストです」と入力して「OK」
3. またステップリストに追加されます

🔰 Step3: 実行してみよう
1. 右上の「⚡実行」ボタンを押すか、F5キーを押す
2. メモ帳が開いて、自動で文字が入力されるのを確認！
3. うまくいかない場合は ESCキーで停止できます

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 画面の見方（どこに何があるの？）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📍 左側：ステップリスト
  → 作成した操作の一覧。上から順番に実行されます。
  → ドラッグで順番を変更、右クリックで編集・削除可能

📍 中央：操作ボタンエリア  
  → よく使う操作のボタンが並んでいます。
  → 📷画像クリック、⌨キー入力、📝文字入力、⏱待機など

📍 右側：プレビューエリア
  → 選択したステップの画像や詳細情報を表示

📍 上部：メニューバー
  → ファイル保存・読込、実行・停止ボタンなど

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 重要！成功のコツ
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✅ 操作の間には必ず「⏱待機」を入れる
   → アプリの起動や画面の切り替わりを待つために1-3秒の待機が重要

✅ 画像認識がうまくいかない時は
   → 類似度を0.8から0.7に下げてみる
   → クリック対象の画像を小さく・明確に撮り直す

✅ 最初は簡単な操作から始める
   → いきなり複雑な作業ではなく、2-3ステップから始める

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ ステップリストの便利操作
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  🖱️ ドラッグ&ドロップ → ステップの順番を変更
  🖱️ ダブルクリック → ステップを編集
  🖱️ 右クリック → コピー・貼り付け・削除メニュー
  ⌨️  Delete → 選択したステップを削除
  ⌨️  Space → ステップの有効/無効を切り替え（一時的にスキップ）
  ⌨️  Ctrl+Z/Y → 操作の取り消し・やり直し

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ ファイルの保存・読込
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💾 保存：Ctrl+S または メニュー「ファイル」→「保存」
📂 読込：Ctrl+O または メニュー「ファイル」→「開く」
📋 作成した自動化スクリプトは.jsonファイルで保存されます

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🚨 困った時は「📖 詳細マニュアル」をチェック！
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        """
        self._show_help_window(help_text, "🚀 クイックスタートガイド", show_detailed_button=True)
    
    def show_help_detailed(self):
        """詳細なマニュアルを表示"""
        help_text = f"""
{AppConfig.APP_NAME} v{AppConfig.VERSION}
📖 詳細操作マニュアル

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 活用シーン別ガイド
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🏢 業務自動化の実例

📌 日次レポート作成
   Excel起動 → データ更新 → PDF保存

📌 メール一括送信  
   顧客リスト → メール作成 → 送信

📌 Webサイト監視
   ページ確認 → 変更検知 → 通知

📌 ファイル整理
   フォルダ作成 → ファイル移動 → 削除

📌 システムバックアップ
   データ圧縮 → サーバー転送 → 確認

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 機能別詳細ガイド
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🔸 画像クリック機能

    【設定項目と推奨値】
    • 信頼度: 0.8～0.9（高精度）/ 0.6～0.7（柔軟）
    • クリックタイプ: single/double/right
    • リトライ回数: 3～5回（通常）/ 10回（不安定時）
    • リトライ間隔: 1～2秒

    使用時の注意：
    • 画像は特徴的な部分（文字、アイコン等）を含める
    • 背景が変化しない要素を選択
    • 画像サイズは50×50～200×200ピクセルが最適

🔸 座標クリック機能

    【特性と詳細】
    • 精度: ピクセル単位の正確な位置指定
    • 依存性: 画面解像度・DPI設定に依存
    • 用途: 固定UI要素、決まった位置の操作
    • 制限事項: ウィンドウ移動時は座標がずれる

🔸 画像オフセットクリック機能

    【パラメータ説明】
    • 基準画像: クリック位置の基準となる画像
    • Xオフセット: 基準から右方向への移動量(px)
    • Yオフセット: 基準から下方向への移動量(px)  
    • 信頼度: 基準画像の認識精度(0.5～1.0)

    応用例：メニュー項目、一覧の特定行、ボタンの隣接要素のクリック

🔸 ドラッグ操作機能

    基本設定：開始座標 → 終了座標 + ドラッグ時間
    用途：ファイル移動、範囲選択、スライダー操作
    時間設定：0.5～2.0秒（操作対象により調整）

🔸 キーボード入力機能

    【対応キー一覧】
    • ファンクション: F1～F12
    • 制御キー: Enter, Tab, Space, Esc など
    • 方向キー: ↑↓←→, Home, End, PgUp, PgDn
    • Ctrl組み合わせ: Ctrl+C/V/A/Z/Y/X/S/O/N/F/H/R
    • Shift組み合わせ: Shift+Tab/F10/Del/Insert  
    • Alt組み合わせ: Alt+Tab/F4/←/→/Enter
    • Win組み合わせ: Win+D/E/R/L/Tab
    • システムキー: Ctrl+Alt+Del, Ctrl+Shift+Esc

    注意：a-z、0-9の文字キーは除外されています（誤操作防止）

🔸 テキスト入力機能

    対応形式：
    • 単一行テキスト（通常の入力フィールド用）
    • 複数行テキスト（テキストエリア、メール本文等）
    • 日本語・多国語対応（Unicode完全対応）
    • 特殊文字・記号（©, ™, ®, €, ¥ 等）

🔸 待機時間機能

    【待機タイプと設定方法】
    • 秒数指定: 0.1～100秒 → 一般的な待機
    • 時刻指定: HH:MM:SS形式 → 定時実行

    推奨待機時間：
    • 画面切り替え後：2～3秒
    • ファイル保存後：1～2秒  
    • ネットワーク処理後：3～5秒

🔸 コマンド実行機能

    対応形式：Windows CMD, PowerShell, Batch
    設定項目：タイムアウト時間、実行完了待機
    セキュリティ：危険なコマンドは実行前に確認

🔸 繰り返し処理機能

    • 指定回数リピート（1～10000回）
    • ネスト対応（繰り返しの中に繰り返し）
    • 無限ループ防止機能内蔵

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 完全ショートカットリファレンス
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【実行・停止】         【編集操作】           【ファイル操作】
F5    実行開始          Ctrl+Z  元に戻す       Ctrl+S  保存
F6    実行停止          Ctrl+Y  やり直し       Ctrl+O  開く
ESC   緊急停止          Ctrl+C  コピー         Ctrl+N  新規作成
                       Ctrl+V  貼り付け       
【表示・検索】         Ctrl+X  切り取り       【システム】
Ctrl+K コマンドパレット Delete  削除           Alt+F4  終了
Ctrl+F 検索            Insert  挿入           F1      ヘルプ
F11   フルスクリーン                          F12     開発者ツール

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ トラブル解決ガイド
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🚨 画像認識エラー対策

    【症状】画像が見つからない、クリック位置がずれる
    【解決策】
    • 信頼度を0.8→0.7→0.6と段階的に下げる
    • 画像を再撮影（より鮮明で特徴的な部分）
    • 画像サイズを調整（推奨：100×100px前後）
    • モニター解像度・拡大率を確認
    • 対象ウィンドウを最前面に表示

🚨 動作不安定対策

    【症状】処理が途中で止まる、エラーが頻発する
    【解決策】
    • 各操作間の待機時間を2～3秒に延長
    • CPUメモリ使用量を確認、他アプリを終了
    • セキュリティソフトの除外設定を追加
    • 管理者権限でアプリケーションを実行
    • Windowsの視覚効果を「パフォーマンス優先」

🚨 キー入力問題対策

    【症状】キーが入力されない、意図と違う文字入力
    【解決策】
    • 対象アプリがアクティブ（最前面）か確認
    • IME状態を確認（日本語/英語入力モード）
    • NumLock/CapsLockの状態を確認
    • キー入力前に0.5～1秒の待機時間を挿入
    • ウィンドウにクリックしてフォーカスを確実に

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 効果的な自動化のコツ
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【設計の基本原則】
1. 小さな単位で分割：複雑な処理は小さなステップに分ける
2. 適切な待機時間：急いで実行せず、確実性を重視
3. エラー対策：代替手順や複数パターンを用意
4. 定期的な検証：環境変化に応じて動作確認

【画像素材の最適化】
• 解像度：100×100px前後が認識精度と速度のバランス良
• 内容：文字やアイコンなど特徴的な要素を含める
• 保存先：専用フォルダで整理、バックアップも取る
• 命名：わかりやすい名前（例：login_button.png）

【パフォーマンス向上】
• 不要なリトライ設定を避ける（処理時間短縮）
• 画像認識の信頼度は必要最低限に設定
• ログレベルを調整してファイル容量を管理
• 定期的に一時ファイルをクリーンアップ

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        """
        self._show_help_window(help_text, "📖 詳細操作マニュアル", show_detailed_button=False)

    def show_help(self):
        """デフォルトのヘルプ表示（簡単なマニュアルを表示）"""
        self.show_help_simple()
    
    def _show_help_window(self, help_text, title, show_detailed_button=False):
        """ヘルプウィンドウを表示する共通メソッド"""
        # スクロール可能なヘルプウィンドウを作成
        help_window = tk.Toplevel(self.root)
        help_window.title(title)
        help_window.geometry("700x800")
        help_window.configure(bg="#2b2b2b")
        help_window.resizable(True, True)
        
        # ウィンドウを中央に配置
        help_window.update_idletasks()
        x = (help_window.winfo_screenwidth() // 2) - (help_window.winfo_width() // 2)
        y = (help_window.winfo_screenheight() // 2) - (help_window.winfo_height() // 2)
        help_window.geometry(f"+{x}+{y}")
        
        # メインフレーム（ダークテーマ統一）
        main_frame = tk.Frame(help_window, bg="#2b2b2b")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # スクロール可能なテキストエリア
        text_frame = tk.Frame(main_frame, bg="#2b2b2b")
        text_frame.pack(fill="both", expand=True)
        
        # テキストウィジェットとスクロールバー
        text_widget = tk.Text(text_frame, 
                             wrap=tk.WORD, 
                             font=("Meiryo UI", 10),
                             bg="#3c3c3c",
                             fg="white",
                             insertbackground="white",
                             selectbackground="#0078d4",
                             relief="flat",
                             padx=15,
                             pady=15)
        
        scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview,
                                bg="#3c3c3c", troughcolor="#2b2b2b", activebackground="#555555")
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        # テキストを挿入
        text_widget.insert("1.0", help_text.strip())
        text_widget.configure(state="disabled")  # 読み取り専用
        
        # 配置
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # ボタンフレーム（ダークテーマ統一）
        button_frame = tk.Frame(main_frame, bg="#2b2b2b")
        button_frame.pack(pady=(15, 0), fill="x")
        
        # 詳細マニュアルボタン（簡単マニュアルの場合のみ表示）
        if show_detailed_button:
            detailed_button = tk.Button(button_frame, text="📖 詳細マニュアル", 
                                        command=lambda: [help_window.destroy(), self.show_help_detailed()],
                                        bg="#0078d4",
                                        fg="white",
                                        font=("Meiryo UI", 10, "bold"),
                                        relief="flat",
                                        cursor="hand2",
                                        padx=20,
                                        pady=8,
                                        width=15)
            detailed_button.pack(side="left")
        
        # 閉じるボタン
        close_button = tk.Button(button_frame, text="閉じる", 
                                command=help_window.destroy,
                                bg="#555555",
                                fg="white",
                                font=("Meiryo UI", 10),
                                relief="flat",
                                cursor="hand2",
                                padx=20,
                                pady=8,
                                width=10)
        close_button.pack(side="right")
        
    def update_status(self, text):
        """新UIと旧UIの両方に対応したステータス更新"""
        if hasattr(self, 'main_status_label'):
            self.main_status_label.config(text=text)
        elif hasattr(self, 'status_label'):
            self.status_label.config(text=text)
            
    def on_monitor_selected(self, event=None):
        if hasattr(self, 'monitor_var'):
            monitor_idx = int(self.monitor_var.get().split()[1])
            self.select_monitor(monitor_idx)

    def setup_header(self, parent):
        """ヘッダーエリアを構築"""
        header_frame = ttk.Frame(parent, style='Modern.TFrame')
        header_frame.pack(fill='x', pady=(0, 10))
        
        # タイトル
        title_label = ttk.Label(header_frame, 
                               text=f"{AppConfig.APP_NAME} v{AppConfig.VERSION}",
                               style='Title.TLabel')
        title_label.pack(side='left')
        
        # 右側のコントロール
        right_controls = ttk.Frame(header_frame, style='Modern.TFrame')
        right_controls.pack(side='right')
        
        # ヘルプボタン（?ボタン）
        help_button = ttk.Button(right_controls, text="?", 
                                command=self.show_help,
                                style='Modern.TButton',
                                width=3)
        help_button.pack(side='right', padx=(5, 0))
        
        # モニター情報
        self.monitor_info_label = ttk.Label(right_controls, 
                                          text="モニター: 未選択", 
                                          style='Modern.TLabel')
        self.monitor_info_label.pack(side='right', padx=(0, 10))
    
    
    # def setup_right_panel(self, parent):  # 旧バージョン - 新UIで置き換え済み
    def setup_right_panel_legacy(self, parent):
        """右パネル（プレビューとコントロール）を構築"""
        right_frame = ttk.Frame(parent, style='Modern.TFrame')
        right_frame.pack(side='right', fill='y', padx=(10, 0))
        right_frame.configure(width=210)  # 右半分のサイズを1/2に縮小
        
        # 画像プレビュー（拡張版）
        self.image_preview = ImagePreviewWidget(right_frame, self)
        self.image_preview.frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # コントロールパネル
        self.setup_control_panel(right_frame)
    
    # def setup_toolbar(self, parent):  # 旧バージョン - 新UIで置き換え済み
    def setup_toolbar_legacy(self, parent):
        """ツールバーを構築"""
        toolbar_frame = ttk.Frame(parent, style='Modern.TFrame')
        toolbar_frame.pack(fill='x', pady=(0, 5))
        
        # ステップ追加ボタン群
        add_frame = ttk.Frame(toolbar_frame, style='Modern.TFrame')
        add_frame.pack(side='left')
        
        buttons = [
            ("📷 画像C", self.add_step_image_click, "画像をクリック"),
            ("🖱️ 座標クリック", self.add_step_coord_click, "座標でクリック"),
            ("🎯 画像O", self.add_step_image_relative_right_click, "画像オフセットクリック"),
            ("⏱️ 待機", self.add_step_sleep, "待機時間"),
            ("⌨️ キー", self.add_step_key_custom, "カスタムキー"),
            ("📝 文字", self.add_step_custom_text, "文字列入力")
        ]
        
        for i, (text, command, tooltip) in enumerate(buttons):
            btn = ttk.Button(add_frame, text=text, command=command, 
                           style='Modern.TButton', width=8)
            btn.grid(row=i//3, column=i%3, padx=2, pady=2)
            self.add_tooltip(btn, tooltip)
        
        # 編集ボタン群
        edit_frame = ttk.Frame(toolbar_frame, style='Modern.TFrame')
        edit_frame.pack(side='right')
        
        edit_buttons = [
            ("✏️ 編集", self.edit_selected_step),
            ("🗑️ 削除", self.delete_selected),
            ("📋 コピー", self.copy_selected_step),
            ("📎 貼付", self.paste_step)
        ]
        
        for text, command in edit_buttons:
            btn = ttk.Button(edit_frame, text=text, command=command, 
                           style='Modern.TButton', width=8)
            btn.pack(side='left', padx=2)
    
    # def setup_control_panel(self, parent):  # 旧バージョン - 新UIで置き換え済み
    def setup_control_panel_legacy(self, parent):
        """コントロールパネルを構築"""
        control_frame = ttk.Frame(parent, style='Modern.TFrame')
        control_frame.pack(fill='x')
        
        # タイトル
        ttk.Label(control_frame, text="コントロール", 
                 style='Title.TLabel').pack(pady=(0, 10))
        
        # 実行ボタン（サイズ縮小）
        self.start_button = ttk.Button(control_frame, text="実行", 
                                     command=self.run_all_steps,
                                     style='Modern.TButton',
                                     width=10)
        self.start_button.pack(pady=3)
        
        self.stop_button = ttk.Button(control_frame, text="停止", 
                                    command=self.stop_execution,
                                    style='Modern.TButton', 
                                    state='disabled',
                                    width=10)
        self.stop_button.pack(pady=3)
        
        # 設定セクション
        settings_section = ttk.LabelFrame(control_frame, text="リストから選択", 
                                        style='Modern.TLabelframe')
        settings_section.pack(fill='x', pady=5)
        
        
        config_frame = ttk.Frame(settings_section, style='Modern.TFrame')
        config_frame.pack(fill='x', padx=10, pady=5)
        
        self.config_var = tk.StringVar(value="デフォルト")
        config_combo = ttk.Combobox(config_frame, textvariable=self.config_var,
                                  values=["デフォルト", "カスタム1", "カスタム2"],
                                  width=15, state="readonly", font=AppConfig.FONTS['small'])
        config_combo.pack(fill='x', expand=True)
        
        # ファイル操作
        file_frame = ttk.LabelFrame(control_frame, text="ファイル", 
                                   style='Modern.TFrame')
        file_frame.pack(pady=5)
        
        file_buttons = [
            ("保存", self.save_config),
            ("読み込み", self.load_config)
        ]
        
        for text, command in file_buttons:
            btn = ttk.Button(file_frame, text=text, command=command,
                           style='Modern.TButton', width=10)
            btn.pack(padx=5, pady=2)
    
    # def setup_status_bar(self, parent):  # 旧バージョン - 新UIで置き換え済み
    def setup_status_bar_legacy(self, parent):
        """ステータスバーを構築"""
        status_frame = ttk.Frame(parent, style='Modern.TFrame')
        status_frame.pack(fill='x', pady=(10, 0))
        
        # ステータスラベル
        self.status_label = ttk.Label(status_frame, text="待機中", 
                                    style='Modern.TLabel')
        self.status_label.pack(side='left')
        
        # 進行状況
        self.progress_var = tk.StringVar(value="")
        self.progress_label = ttk.Label(status_frame, textvariable=self.progress_var,
                                       style='Modern.TLabel')
        self.progress_label.pack(side='right')
    
    def setup_hotkeys(self):
        """拡張されたホットキーを設定"""
        # 基本操作
        self.root.bind('<Control-s>', lambda e: self.save_config())
        self.root.bind('<Control-o>', lambda e: self.load_config())
        self.root.bind('<Control-n>', lambda e: self.clear_all_steps())
        
        # 実行制御
        self.root.bind('<F5>', lambda e: self.run_all_steps())
        self.root.bind('<F6>', lambda e: self.stop_execution())
        self.root.bind('<F7>', lambda e: self.run_from_selected())
        self.root.bind('<F8>', lambda e: self.run_from_failed())
        
        # ESC緊急停止
        self.root.bind('<Escape>', lambda e: self.emergency_stop())
        self.root.focus_set()  # フォーカスを確保してキーバインドを有効化
        
        # 編集操作
        self.root.bind('<Delete>', lambda e: self.delete_selected())
        self.root.bind('<Control-c>', lambda e: self.copy_selected_step())
        self.root.bind('<Control-v>', lambda e: self.paste_step())
        
        # Undo/Redo
        self.root.bind('<Control-z>', lambda e: self.undo())
        self.root.bind('<Control-y>', lambda e: self.redo())
        
        
        # キャプチャ
        self.root.bind('<Control-Shift-s>', lambda e: self.image_preview.start_capture())
        # スクリーンショット撮影
        self.root.bind('<Control-Shift-F12>', lambda e: self.add_step_screenshot())
        
        # デバッグ用: キーバインドテスト
        self.root.bind('<Control-Shift-t>', lambda e: self.test_key_binding())


    def setup_button_frames_old(self, parent: ttk.Frame):
        # 左: ステップ追加
        left_frame = ttk.Frame(parent)
        left_frame.pack(side=tk.LEFT, padx=5, fill=tk.Y)
        ttk.Label(left_frame, text="ステップ追加", font=("Arial", 12, "bold")).pack(pady=5)
        button_configs = [
            # 1列目
            (0, 0, "画像クリック", self.add_step_image_click, "画像をクリック/ダブルクリックするステップを追加"),
            (1, 0, "座標クリック", self.add_step_coord_click, "指定座標でクリックするステップを追加"),
            (2, 0, "画像オフセットクリック", self.add_step_image_relative_right_click, "画像の位置からオフセットした位置をクリックするステップを追加"),
            (3, 0, "スリープ", self.add_step_sleep, "待機時間を追加"),
            # 2列目（コピーとペーストを←→と入れ替え）
            (0, 1, "←", lambda: self.add_step_key_action("key", "left"), "左キーを追加"),
            (1, 1, "→", lambda: self.add_step_key_action("key", "right"), "右キーを追加"),
            (2, 1, "↑", lambda: self.add_step_key_action("key", "up"), "上キーを追加"),
            (3, 1, "↓", lambda: self.add_step_key_action("key", "down"), "下キーを追加"),
            # 3列目（←→をコピーとペーストと入れ替え）
            (0, 2, "コピー", lambda: self.add_step_key_action("copy", "ctrl+c"), "コピー操作を追加"),
            (1, 2, "ペースト", lambda: self.add_step_key_action("paste", "ctrl+v"), "ペースト操作を追加"),
            (2, 2, "Tab", lambda: self.add_step_key_action("key", "tab"), "Tabキーを追加"),
            (3, 2, "Enter", lambda: self.add_step_key_action("key", "enter"), "Enterキーを追加"),
            # 4列目
            (0, 3, "カスタムキー", self.add_step_key_custom, "任意のキー操作を追加"),
            (1, 3, "カスタム文字列", self.add_step_custom_text, "任意の文字列を入力するステップを追加"),
        ]
        button_grid = ttk.Frame(left_frame)
        button_grid.pack(fill=tk.X)
        for row, col, text, command, tooltip in button_configs:
            btn = ttk.Button(button_grid, text=text, command=command)
            btn.grid(row=row, column=col, padx=5, pady=5, sticky="ew")
            self.add_tooltip(btn, tooltip)
        for i in range(4):
            button_grid.columnconfigure(i, weight=1)

        # 中央: ステップ操作
        center_frame = ttk.Frame(parent)
        center_frame.pack(side=tk.LEFT, padx=5, fill=tk.Y)
        ttk.Label(center_frame, text="ステップ操作", font=("Arial", 12, "bold")).pack(pady=5)
        for text, command, tooltip in [
            ("削除", self.delete_selected, "選択したステップを削除"),
            ("編集", self.edit_selected_step, "選択したステップを編集"),
            ("全クリア", self.clear_all_steps, "すべてのステップをクリア"),
        ]:
            btn = ttk.Button(center_frame, text=text, command=command)
            btn.pack(fill=tk.X, pady=5)
            self.add_tooltip(btn, tooltip)

        # 右: 設定・実行
        right_frame = ttk.Frame(parent)
        right_frame.pack(side=tk.RIGHT, padx=5, fill=tk.Y)
        ttk.Label(right_frame, text="設定・実行", font=("Arial", 12, "bold")).pack(pady=5)
        # self.start_button = ttk.Button(right_frame, text="実行開始", command=self.run_all_steps)
        # self.start_button.pack(fill=tk.X, pady=5)
        # self.add_tooltip(self.start_button, "自動操作を開始")
        # self.stop_button = ttk.Button(right_frame, text="実行停止", command=self.stop_execution, state="disabled")
        # self.stop_button.pack(fill=tk.X, pady=5)
        # self.add_tooltip(self.stop_button, "自動操作を停止")
        
        for text, command, tooltip in [
            ("保存", self.save_config, "設定をファイルに保存"),
            ("読込", self.load_config, "設定をファイルから読み込み"),
            # ("全モニター検索", self.run_all_monitors, "全モニターで画像を検索"),
        ]:
            btn = ttk.Button(right_frame, text=text, command=command)
            btn.pack(fill=tk.X, pady=5)
            self.add_tooltip(btn, tooltip)

        # ループ回数とモニター選択
        settings_frame = ttk.Frame(right_frame)
        settings_frame.pack(fill=tk.X, pady=5)
        ttk.Label(settings_frame, text="モニター選択").pack(side=tk.LEFT, padx=5)
        self.monitor_var = tk.StringVar(value="0")
        monitor_menu = ttk.OptionMenu(settings_frame, self.monitor_var, "0", *[str(i) for i in range(len(self.monitors))], command=self.select_monitor)
        monitor_menu.pack(side=tk.LEFT, padx=5)
        self.add_tooltip(monitor_menu, "操作対象のモニターを選択")

    def delete_selected(self):
        selected = self.tree.selection()
        if selected:
            # 状態を保存（Undo用）
            self.save_state(f"ステップ削除: {len(selected)}個")
            
            for item in selected:
                index = self.tree.index(item)
                self.tree.delete(item)
                self.steps.pop(index)
            
            # ステップが削除されたら設定コンボボックスをクリア（編集中状態を示す）
            if hasattr(self, 'config_combo'):
                self.config_combo.set("")
            
            logger.info("選択したステップを削除")

    def clear_all_steps(self):
        if messagebox.askyesno("確認", "すべてのステップを削除しますか？"):
            # 状態を保存（Undo用）
            self.save_state("全ステップクリア")
            
            self.steps.clear()
            self.refresh_tree()
            
            # 設定コンボボックスをクリア（新規状態を示す）
            if hasattr(self, 'config_combo'):
                self.config_combo.set("")
            
            logger.info("すべてのステップをクリア")


    def add_tooltip(self, widget: tk.Widget, text: str):
        tooltip_window = None

        def enter(event):
            nonlocal tooltip_window
            if tooltip_window:
                return
                
            # Toplevelウィンドウでツールチップを作成
            tooltip_window = tk.Toplevel(widget)
            tooltip_window.wm_overrideredirect(True)
            tooltip_window.configure(background="#ffffe0")
            
            # ラベルを作成
            label = tk.Label(tooltip_window, text=text, 
                           background="#ffffe0", foreground="black",
                           relief="solid", borderwidth=1, 
                           font=("Arial", 8), padx=4, pady=2)
            label.pack()
            
            # 位置を計算
            x = widget.winfo_rootx() + 5
            y = widget.winfo_rooty() + widget.winfo_height() + 5
            
            # サイズを取得して画面境界チェック
            tooltip_window.update_idletasks()
            tooltip_width = tooltip_window.winfo_reqwidth()
            tooltip_height = tooltip_window.winfo_reqheight()
            
            screen_width = tooltip_window.winfo_screenwidth()
            screen_height = tooltip_window.winfo_screenheight()
            
            # 画面からはみ出ないよう調整
            if x + tooltip_width > screen_width:
                x = screen_width - tooltip_width - 10
            if y + tooltip_height > screen_height:
                y = widget.winfo_rooty() - tooltip_height - 5
            
            tooltip_window.geometry(f"+{x}+{y}")

        def leave(event):
            nonlocal tooltip_window
            if tooltip_window:
                tooltip_window.destroy()
                tooltip_window = None

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def select_monitor(self, value: str):
        self.selected_monitor = int(value)
        monitor = self.monitors[self.selected_monitor]
        scale_warning = " (スケーリングが異なる可能性)" if self.dpi_scale != 100.0 else ""
        # モニター情報を更新（新UIと旧UIの両方をサポート）
        monitor_text = f"モニター {self.selected_monitor}: {monitor.width}x{monitor.height} at ({monitor.x}, {monitor.y}){scale_warning}"
        
        if hasattr(self, 'system_info_label'):
            self.system_info_label.config(text=monitor_text)
        elif hasattr(self, 'monitor_info_label'):
            self.monitor_info_label.config(text=monitor_text)
        logger.info(f"モニター選択: {self.selected_monitor}")
        self.save_last_config()

    def save_last_config(self):
        try:
            config = {"last_monitor": self.selected_monitor, "loop_count": self.loop_count}
            if os.path.exists(self.config_file):
                with open(self.config_file, "r", encoding="utf-8") as f:
                    config.update(json.load(f))
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            logger.info(f"最終設定保存: モニター={self.selected_monitor}, ループ回数={self.loop_count}")
        except Exception as e:
            logger.error(f"最終設定保存エラー: {e}")

    def add_step(self, step: Step):
        """新しいステップを追加（選択行の下に挿入）"""
        if not step.validate():
            raise ValueError("無効なステップです")
        
        # 状態を保存（Undo用）
        self.save_state(f"ステップ追加: {step.type}")
        
        # 選択されている行のインデックスを取得
        selected_items = self.tree.selection()
        if selected_items:
            # 選択されている行がある場合、その下に挿入
            selected_item = selected_items[0]
            selected_index = self.tree.index(selected_item)
            insert_index = selected_index + 1
        else:
            # 選択されている行がない場合、末尾に追加
            insert_index = len(self.steps)
        
        self.steps.insert(insert_index, step)
        self.refresh_tree()
        
        # 新しいステップが追加されたら設定コンボボックスをクリア（編集中状態を示す）
        if hasattr(self, 'config_combo'):
            self.config_combo.set("")
        
        # 新しく挿入されたステップを選択
        if self.steps:
            tree_children = self.tree.get_children()
            if insert_index < len(tree_children):
                new_item = tree_children[insert_index]
                self.tree.selection_set(new_item)
                self.tree.see(new_item)
        
        # 自動保存機能
        if self.auto_save_enabled:
            self.save_last_config()
        
        logger.info(f"ステップ追加: {step.type} - {step.comment}")
        
        # 自動保存機能（既にadd_stepで状態保存されている）
    
    def move_step(self, from_index: int, to_index: int):
        """ステップの順序を変更"""
        if 0 <= from_index < len(self.steps) and 0 <= to_index < len(self.steps):
            # 状態を保存（Undo用）
            self.save_state(f"ステップ移動: {from_index} → {to_index}")
            
            step = self.steps.pop(from_index)
            self.steps.insert(to_index, step)
            self.refresh_tree()
            
            # 移動後のアイテムを選択
            moved_item = self.tree.get_children()[to_index]
            self.tree.selection_set(moved_item)
            self.tree.see(moved_item)
            
            logger.info(f"ステップ移動: {from_index} -> {to_index}")
    
    def copy_selected_step(self):
        """選択したステップをコピー"""
        selection = self.tree.selection()
        if selection:
            index = self.tree.index(selection[0])
            if index < len(self.steps):
                self.clipboard_step = self.steps[index].to_dict()
                status_text = f"📋 ステップをコピー: {self.steps[index].comment}"
                if hasattr(self, 'main_status_label'):
                    self.main_status_label.configure(text=status_text)
                elif hasattr(self, 'status_label'):
                    self.status_label.configure(text=status_text)
                logger.info(f"ステップコピー: {self.steps[index].comment}")
    
    def paste_step(self):
        """コピーしたステップを貼り付け"""
        if self.clipboard_step:
            try:
                new_step = Step.from_dict(self.clipboard_step.copy())
                new_step.comment += " (コピー)"
                self.add_step(new_step)
                status_text = f"📄 ステップを貼り付け: {new_step.comment}"
                if hasattr(self, 'main_status_label'):
                    self.main_status_label.configure(text=status_text)
                elif hasattr(self, 'status_label'):
                    self.status_label.configure(text=status_text)
                logger.info(f"ステップ貼り付け: {new_step.comment}")
            except Exception as e:
                self.show_error_with_sound("エラー", f"ステップの貼り付けに失敗しました: {e}")
                logger.error(f"ステップ貼り付けエラー: {e}")
        else:
            messagebox.showwarning("警告", "コピーされたステップがありません")
    
    def toggle_step_enabled(self, index: int):
        """ステップの有効/無効を切り替え"""
        if 0 <= index < len(self.steps):
            # 状態を保存（Undo用）
            status = "無効化" if self.steps[index].enabled else "有効化"
            self.save_state(f"ステップ{status}: {index}")
            
            self.steps[index].enabled = not self.steps[index].enabled
            self.refresh_tree()
            
            status = "有効" if self.steps[index].enabled else "無効"
            # ステータス更新（新UIと旧UIの両方をサポート）
            status_text = f"ステップを{status}に変更: {self.steps[index].comment}"
            
            if hasattr(self, 'main_status_label'):
                self.main_status_label.configure(text=status_text)
            elif hasattr(self, 'status_label'):
                self.status_label.configure(text=status_text)
            logger.info(f"ステップ{status}化: {self.steps[index].comment}")

    def get_type_display(self, step: Step) -> str:
        """統一アイコンを使用したアクション表示"""
        return AppConfig.get_step_display_name(step.type)

    def get_params_display(self, step: Step) -> str:
        """簡潔なパラメータ表示"""
        if step.type == "image_click":
            filename = os.path.basename(step.params['path'])
            click_type_map = {"double": "ダブル", "right": "右", "single": "シングル"}
            click_type = click_type_map.get(step.params.get('click_type', 'single'), "シングル")
            return f"{filename} • {click_type} • 信頼度:{step.params['threshold']}"
        elif step.type == "coord_click":
            click_type_map = {"double": "ダブル", "right": "右", "single": "シングル"}
            click_type = click_type_map.get(step.params.get('click_type', 'single'), "シングル")
            return f"({step.params['x']}, {step.params['y']}) • {click_type}"
        elif step.type == "coord_drag":
            return f"({step.params['start_x']}, {step.params['start_y']}) → ({step.params['end_x']}, {step.params['end_y']}) • {step.params['duration']}秒"
        elif step.type == "image_relative_right_click":
            filename = os.path.basename(step.params['path'])
            click_type_map = {"double": "ダブル", "right": "右", "single": "シングル"}
            click_type = click_type_map.get(step.params.get('click_type', 'right'), "右")
            return f"{filename} • {click_type} • オフセット:({step.params['offset_x']}, {step.params['offset_y']})"
        elif step.type == "sleep":
            wait_type = step.params.get("wait_type", "sleep")
            if wait_type == "scheduled":
                return f"時刻指定: {step.params['scheduled_time']}"
            else:
                return f"{step.params.get('seconds', 1.0)}秒間"
        elif step.type in ["copy", "paste", "key"]:
            return f"[{step.params['key']}]"
        elif step.type == "custom_text":
            text = step.params['text'].replace('\n', ' ').replace('\r', ' ')[:30]
            text = text + ('...' if len(step.params['text']) > 30 else '')
            return f'"{text}"'
        elif step.type == "cmd_command":
            command = step.params['command'].replace('\n', ' ').replace('\r', ' ')[:50]
            command = command + ('...' if len(step.params['command']) > 50 else '')
            return f'[CMD] {command}'
        elif step.type == "repeat_start":
            return f"{step.params['count']}回繰り返し開始"
        elif step.type == "repeat_end":
            return "繰り返し終了"
        return str(step.params)
    
    def get_clean_comment(self, step: Step) -> str:
        """デフォルトコメントを除去してユーザーコメントのみ表示"""
        # デフォルトコメントパターン
        default_patterns = [
            "画像クリック:",
            "座標右クリック",
            "画像オフセットクリック:",
            "待機時間",
            "キー操作",
            "カスタム文字列入力",
            "キー入力間隔",
            "右クリック間隔",
            "画像クリック後待機"
        ]
        
        comment = step.comment.strip()
        
        # デフォルトパターンをチェック
        for pattern in default_patterns:
            if pattern in comment:
                # パターンで分割して、パターン以降の部分をチェック
                parts = comment.split(pattern, 1)
                if len(parts) > 1:
                    remaining = parts[1].strip()
                    # ファイル名部分を除去してユーザーコメントを抽出
                    if pattern == "画像クリック:":
                        # "画像クリック: filename.png" のfilename.png部分を除去
                        # 改行、" - "、またはファイル名後の最初の空白以降をユーザーコメントとみなす
                        if "\n" in remaining:
                            user_comment = remaining.split("\n", 1)[1].strip()
                            if user_comment:
                                return user_comment
                        elif " - " in remaining:
                            user_comment = remaining.split(" - ", 1)[1].strip()  
                            if user_comment:
                                return user_comment
                        # その他の場合は、ファイル名のみとみなして空を返す
                        return ""
                    else:
                        # 他のパターンの場合、パターン後に何かあればそれを返す
                        if remaining:
                            return remaining
                # パターンのみの場合は空文字列
                return ""
        
        # デフォルトパターンが含まれない場合はそのまま返す
        return comment

    def add_step_image_click(self):
        """画像クリック/ダブルクリックのステップを追加（クリップボード画像対応）"""
        # 拡張されたダイアログを表示
        dialog = EnhancedImageDialog(self.root, "画像クリック設定")
        image_path = dialog.get_image_path()
        
        if not image_path:
            return

        fields = [
            {"key": "threshold", "label": "信頼度(0.5-1.0):", "type": "float", "default": "0.8", "min": 0.5, "max": 1.0},
            {"key": "click_type", "label": "クリックタイプ:", "type": "combobox", "default": "single", "values": AppConfig.CLICK_TYPES, "required": True},
            {"key": "retry", "label": "リトライ回数:", "type": "int", "default": "3", "min": 0},
            {"key": "delay", "label": "リトライ間隔(秒):", "type": "float", "default": "1.0", "min": 0.1},
            {"key": "comment", "label": "メモ:", "type": "text", "default": f"画像クリック: {os.path.basename(image_path)}"},
        ]
        dialog = ModernDialog(self.root, "画像クリック設定", fields, width=700, height=800)
        try:
            result = dialog.get_result()
            if result:
                params = {
                    "path": image_path,
                    "threshold": result["threshold"],
                    "click_type": result["click_type"],
                    "retry": result["retry"],
                    "delay": result["delay"],
                }
                self.add_step(Step("image_click", params=params, comment=result["comment"]))
                self.add_step(Step("sleep", params={"seconds": 0.5}, comment="画像クリック後待機"))
        except Exception as e:
            logger.error(f"画像クリック追加エラー: {e}")
            self.show_error_with_sound("エラー", f"画像クリックの追加に失敗しました: {e}")

    def add_step_coord_click(self):
        """座標指定でクリックのステップを追加（リアルタイム座標表示対応）"""
        # 座標選択ダイアログを表示
        coord_dialog = MouseCoordinateDialog(self.root, "座標クリック設定")
        coordinates = coord_dialog.get_coordinates()
        
        if coordinates is None:
            return
            
        fields = [
            {"key": "x", "label": "X座標:", "type": "int", "default": str(coordinates[0])},
            {"key": "y", "label": "Y座標:", "type": "int", "default": str(coordinates[1])},
            {"key": "click_type", "label": "クリックタイプ:", "type": "combobox", "default": "single", "values": AppConfig.CLICK_TYPES, "required": True},
            {"key": "comment", "label": "メモ:", "type": "text", "default": f"座標クリック({coordinates[0]}, {coordinates[1]})"},
        ]
        dialog = ModernDialog(self.root, "座標クリック設定", fields, width=700, height=800)
        try:
            result = dialog.get_result()
            if result:
                params = {"x": result["x"], "y": result["y"], "click_type": result["click_type"]}
                self.add_step(Step("coord_click", params=params, comment=result["comment"]))
                self.add_step(Step("sleep", params={"seconds": 0.5}, comment="クリック間隔"))
        except Exception as e:
            logger.error(f"座標クリック追加エラー: {e}")
            self.show_error_with_sound("エラー", f"座標クリックの設定に失敗しました: {e}")

    def add_step_coord_drag(self):
        """座標間ドラッグのステップを追加"""
        # 開始座標選択
        start_dialog = MouseCoordinateDialog(self.root, "ドラッグ開始座標設定")
        start_coordinates = start_dialog.get_coordinates()
        
        if start_coordinates is None:
            return
            
        # 終了座標選択
        end_dialog = MouseCoordinateDialog(self.root, "ドラッグ終了座標設定")
        end_coordinates = end_dialog.get_coordinates()
        
        if end_coordinates is None:
            return
            
        fields = [
            {"key": "start_x", "label": "開始X座標:", "type": "int", "default": str(start_coordinates[0])},
            {"key": "start_y", "label": "開始Y座標:", "type": "int", "default": str(start_coordinates[1])},
            {"key": "end_x", "label": "終了X座標:", "type": "int", "default": str(end_coordinates[0])},
            {"key": "end_y", "label": "終了Y座標:", "type": "int", "default": str(end_coordinates[1])},
            {"key": "duration", "label": "ドラッグ時間(秒):", "type": "float", "default": "0.5", "min": 0.1, "help": "ドラッグにかける時間"},
            {"key": "comment", "label": "メモ:", "type": "text", "default": f"座標ドラッグ({start_coordinates[0]},{start_coordinates[1]})→({end_coordinates[0]},{end_coordinates[1]})"},
        ]
        dialog = ModernDialog(self.root, "座標ドラッグ設定", fields, width=700, height=800)
        try:
            result = dialog.get_result()
            if result:
                params = {
                    "start_x": result["start_x"], "start_y": result["start_y"],
                    "end_x": result["end_x"], "end_y": result["end_y"],
                    "duration": result["duration"]
                }
                self.add_step(Step("coord_drag", params=params, comment=result["comment"]))
                self.add_step(Step("sleep", params={"seconds": 0.5}, comment="ドラッグ後待機"))
        except Exception as e:
            logger.error(f"座標ドラッグ追加エラー: {e}")
            self.show_error_with_sound("エラー", f"座標ドラッグの設定に失敗しました: {e}")

    def add_step_sleep(self):
        """スリープタイムを追加"""
        fields = [
            {"key": "wait_type", "label": "待機タイプ:", "type": "combobox", "values": ["スリープ(秒数指定)", "時刻指定"], "default": "スリープ(秒数指定)", "on_change": True},
            {"key": "seconds", "label": "待ち時間(秒):", "type": "float", "default": "1.0", "min": 0.1, "show_condition": {"field": "wait_type", "value": "スリープ(秒数指定)"}},
            {"key": "scheduled_time", "label": "実行時刻(HH:MM:SS):", "type": "str", "default": "", "help": "日次実行時刻を指定（例：14:30:00）", "show_condition": {"field": "wait_type", "value": "時刻指定"}},
            {"key": "comment", "label": "メモ:", "type": "text", "default": "待機時間"},
        ]
        dialog = ModernDialog(self.root, "待機時間設定", fields, width=700, height=800)
        try:
            result = dialog.get_result()
            if result:
                params = {}
                if result["wait_type"] == "時刻指定":
                    if not result["scheduled_time"].strip():
                        self.show_error_with_sound("エラー", "時刻指定を選択した場合は実行時刻を入力してください。")
                        return
                    # HH:MM:SS形式の検証
                    try:
                        time_parts = result["scheduled_time"].split(':')
                        if len(time_parts) != 3:
                            raise ValueError("時刻の形式が正しくありません")
                        int(time_parts[0])  # 時
                        int(time_parts[1])  # 分  
                        int(time_parts[2])  # 秒
                        params["scheduled_time"] = result["scheduled_time"]
                        params["wait_type"] = "scheduled"
                    except ValueError:
                        self.show_error_with_sound("エラー", "時刻の形式が正しくありません。HH:MM:SS形式で入力してください。（例：14:30:00）")
                        return
                else:
                    # デフォルトはスリープ（秒数指定）
                    params["seconds"] = result["seconds"]
                    params["wait_type"] = "sleep"
                
                self.add_step(Step("sleep", params=params, comment=result["comment"]))
        except Exception as e:
            logger.error(f"スリープ追加エラー: {e}")
            self.show_error_with_sound("エラー", f"待機時間の設定に失敗しました: {e}")

    def add_step_sleep_seconds(self):
        """秒数指定での待機ステップを追加"""
        fields = [
            {"key": "seconds", "label": "待ち時間(秒):", "type": "float", "default": "1.0", "min": 0.1},
            {"key": "comment", "label": "メモ:", "type": "text", "default": "待機(秒数指定)"},
        ]
        dialog = ModernDialog(self.root, "待機(秒数指定)設定", fields, width=700, height=800)
        try:
            result = dialog.get_result()
            if result:
                params = {"wait_type": "sleep", "seconds": float(result["seconds"])}
                self.add_step(Step("sleep", params=params, comment=result["comment"]))
        except Exception as e:
            logger.error(f"秒数指定スリープ追加エラー: {e}")
            self.show_error_with_sound("エラー", f"秒数指定待機の設定に失敗しました: {e}")

    def add_step_sleep_time(self):
        """時刻指定での待機ステップを追加"""
        from datetime import datetime, timedelta
        
        # デフォルト時刻を現在時刻の1分後に設定
        now = datetime.now()
        next_minute = (now + timedelta(minutes=1)).replace(second=0, microsecond=0)
        default_time = next_minute.strftime("%H:%M:%S")
        
        fields = [
            {"key": "scheduled_time", "label": "実行時刻(HH:MM:SS):", "type": "str", "default": default_time, "help": "実行時刻を指定（例：14:30:00）"},
            {"key": "comment", "label": "メモ:", "type": "text", "default": f"待機(時刻指定: {default_time})"},
        ]
        dialog = ModernDialog(self.root, "待機(時刻指定)設定", fields, width=700, height=800)
        try:
            result = dialog.get_result()
            if result:
                if not result["scheduled_time"].strip():
                    self.show_error_with_sound("エラー", "実行時刻を入力してください。")
                    return
                params = {"wait_type": "scheduled", "scheduled_time": result["scheduled_time"]}
                self.add_step(Step("sleep", params=params, comment=result["comment"]))
        except Exception as e:
            logger.error(f"時刻指定スリープ追加エラー: {e}")
            self.show_error_with_sound("エラー", f"時刻指定待機の設定に失敗しました: {e}")

    def add_step_key_action(self, step_type: str, key: str):
        """キー操作を追加（ダイアログ付き）"""
        fields = [
            {"key": "comment", "label": "メモ:", "type": "text", "default": f"{key} キー操作"},
        ]
        dialog = ModernDialog(self.root, f"{key}設定", fields, width=700, height=800)
        try:
            result = dialog.get_result()
            if result:
                # 統合メソッドを使用
                self.add_specific_key_step(key, result["comment"])
        except Exception as e:
            logger.error(f"キー操作追加エラー: {e}")
            self.show_error_with_sound("エラー", f"キー操作の設定に失敗しました: {e}")


    def add_step_key_custom(self):
        """カスタムキー入力"""
        fields = [
            {"key": "key", "label": "キー (例: Ctrl+C, F5):", "type": "combobox", "default": "enter", "values": AppConfig.KEY_OPTIONS, "required": True},
            {"key": "comment", "label": "メモ:", "type": "text", "default": "カスタムキー入力"},
        ]
        dialog = ModernDialog(self.root, "カスタムキー設定", fields, width=700, height=800)
        try:
            result = dialog.get_result()
            if result:
                # 統合メソッドを使用
                comment = result["comment"]
                self.add_specific_key_step(result["key"], comment)
        except Exception as e:
            logger.error(f"カスタムキー追加エラー: {e}")
            self.show_error_with_sound("エラー", f"カスタムキーの設定に失敗しました: {e}")





    def add_step_custom_text(self):
        """カスタム文字列入力のステップを追加"""
        fields = [
            {"key": "text", "label": "入力文字列:", "type": "text", "height": 8, 
             "default": "", 
             "required": True, "help": "複数行のテキストを記述できます"},
            {"key": "comment", "label": "メモ:", "type": "text", "default": "カスタム文字列入力"},
        ]
        dialog = ModernDialog(self.root, "文字列入力設定", fields, width=700, height=800)
        try:
            result = dialog.get_result()
            if result:
                self.add_step(Step("custom_text", params={"text": result["text"]}, comment=result["comment"]))
                self.add_step(Step("sleep", params={"seconds": 0.5}, comment="文字列入力後待機"))
        except Exception as e:
            logger.error(f"カスタム文字列入力エラー: {e}")
            self.show_error_with_sound("エラー", f"文字列入力の設定に失敗しました: {e}")

    def add_step_repeat_start(self):
        """繰り返し開始ステップを追加"""
        fields = [
            {"key": "count", "label": "繰り返し回数:", "type": "int", "default": "2", "min": 1, "required": True},
            {"key": "comment", "label": "メモ:", "type": "text", "default": "繰り返し開始"},
        ]
        dialog = ModernDialog(self.root, "繰り返し開始設定", fields, width=700, height=800)
        try:
            result = dialog.get_result()
            if result:
                self.add_step(Step("repeat_start", params={"count": result["count"]}, comment=result["comment"]))
        except Exception as e:
            logger.error(f"繰り返し開始追加エラー: {e}")
            self.show_error_with_sound("エラー", f"繰り返し開始の設定に失敗しました: {e}")

    def add_step_repeat_end(self):
        """繰り返し終了ステップを追加"""
        fields = [
            {"key": "comment", "label": "メモ:", "type": "text", "default": "繰り返し終了"},
        ]
        dialog = ModernDialog(self.root, "繰り返し終了設定", fields, width=700, height=800)
        try:
            result = dialog.get_result()
            if result:
                self.add_step(Step("repeat_end", params={}, comment=result["comment"]))
        except Exception as e:
            logger.error(f"繰り返し終了追加エラー: {e}")
            self.show_error_with_sound("エラー", f"繰り返し終了の設定に失敗しました: {e}")

    def add_step_cmd_command(self):
        """cmdコマンド実行ステップを追加"""
        fields = [
            {"key": "command", "label": "コマンド:", "type": "text", "height": 12, 
             "default": "例:\npython script.py\n\nまたは:\nstart chrome https://google.com\n\nまたは:\ncd C:\\MyProject\npython main.py", 
             "required": True, "help": "複数行のコマンドを記述できます"},
            {"key": "timeout", "label": "タイムアウト(秒):", "type": "int", "default": "30", "min": 1, "required": True},
            {"key": "wait_completion", "label": "完了を待つ:", "type": "combobox", "values": ["待つ", "待たない"], "default": "待つ"},
            {"key": "comment", "label": "メモ:", "type": "text", "default": "コマンド実行"},
        ]
        dialog = ModernDialog(self.root, "コマンド実行設定", fields, width=700, height=800)
        try:
            result = dialog.get_result()
            if result:
                params = {
                    "command": result["command"],
                    "timeout": result["timeout"],
                    "wait_completion": result["wait_completion"] == "待つ"  # 文字列からboolに変換
                }
                
                self.add_step(Step("cmd_command", params=params, comment=result["comment"]))
        except Exception as e:
            logger.error(f"コマンド実行追加エラー: {e}")
            self.show_error_with_sound("エラー", f"コマンド実行の設定に失敗しました: {e}")



    def add_step_image_relative_right_click(self):
        """画像指定のオフセット座標でクリックのステップを追加（拡張対応）"""
        # 拡張された画像選択ダイアログ
        dialog = EnhancedImageDialog(self.root, "画像オフセットクリック設定")
        image_path = dialog.get_image_path()
        
        if not image_path:
            return

        fields = [
            {"key": "threshold", "label": "信頼度(0.5-1.0):", "type": "float", "default": "0.8", "min": 0.5, "max": 1.0},
            {"key": "click_type", "label": "クリックタイプ:", "type": "combobox", "default": "single", "values": AppConfig.CLICK_TYPES, "required": True},
            {"key": "offset_x", "label": "Xオフセット:", "type": "int", "default": "0"},
            {"key": "offset_y", "label": "Yオフセット:", "type": "int", "default": "0"},
            {"key": "retry", "label": "リトライ回数:", "type": "int", "default": "3", "min": 1},
            {"key": "delay", "label": "リトライ間隔(秒):", "type": "int", "default": "1", "min": 1},
            {"key": "comment", "label": "メモ:", "type": "text", "default": f"画像オフセットクリック: {os.path.basename(image_path)}"},
        ]
        dialog = ModernDialog(self.root, "画像オフセットクリックの設定", fields, width=700, height=800)
        try:
            result = dialog.get_result()
            if result:
                params = {
                    "path": image_path,
                    "threshold": result["threshold"],
                    "click_type": result["click_type"],
                    "offset_x": result["offset_x"],
                    "offset_y": result["offset_y"],
                    "retry": result["retry"],
                    "delay": result["delay"],
                }
                self.add_step(Step("image_relative_right_click", params=params, comment=result["comment"]))
                self.add_step(Step("sleep", params={"seconds": 1}, comment="オフセットクリック後の待機"))
        except Exception as e:
            logger.error(f"画像オフセットクリック追加エラー: {e}")
            self.show_error_with_sound("エラー", f"画像オフセットクリックの設定に失敗しました: {e}")




    def delete_clicked(self, event=None):
        """選択したステップを削除"""
        try:
            selected_items = self.tree.selection()
            if not selected_items:
                messagebox.showwarning("警告", "削除するステップを選択してください")
                return
            for item in selected_items:
                index = self.tree.index(item) + 1  # 1-based index
                self.tree.delete(item)
                self.steps.pop(index - 1)
                logger.info(f"選択されたステップを削除しました: 行番号={index}")
        except Exception as e:
            logger.error(f"ステップ削除エラー: {e}")
            self.show_error_with_sound("エラー", f"ステップの削除に失敗しました: {e}")

    def clear_all(self):
        """すべてのステップを削除"""
        try:
            if messagebox.askyesno("確認", "すべてのステップを削除しますか？"):
                self.steps.clear()
                self.tree.delete(*self.tree.get_children())
                logger.info("すべてのステップを削除しました")
        except Exception as e:
            logger.error(f"全ステップ削除エラー: {e}")
            self.show_error_with_sound("エラー", f"すべてのステップの削除に失敗しました: {e}")


    def edit_selected_step(self, event=None):
        """選択されたステップを編集"""
        try:
            selected_items = self._validate_selection()
            if not selected_items:
                return
            index = self.tree.index(selected_items[0]) + 1  # 1-based index
            step = self.steps[index - 1]
            fields = []

            if step.type == "image_click":
                fields = [
                    {
                        "key": "threshold",
                        "label": "信頼度（0.5-1.0):",
                        "type": "float",
                        "default": str(step.params["threshold"]),
                        "min": 0.5,
                        "max": 1.0
                    },
                    {
                        "key": "click_type",
                        "label": "Click Type:",
                        "type": "combobox",
                        "default": step.params.get("click_type", "single"),
                        "values": AppConfig.CLICK_TYPES,
                        "required": True
                    },
                    {
                        "key": "retry",
                        "label": "リトライ回数（0-10）:",
                        "type": "int",
                        "default": str(step.params["retry"]),
                        "min": 0,
                    },
                    {
                        "key": "delay",
                        "label": "リトライ間隔(秒):",
                        "type": "float",
                        "default": str(step.params["delay"]),
                        "min": 0.1,
                    },
                    {"key": "comment", "label": "メモ:", "type": "text", "default": step.comment},
                ]
            elif step.type == "coord_click":
                fields = [
                    {"key": "x", "label": "X座標:", "type": "int", "default": str(step.params["x"])},
                    {"key": "y", "label": "Y座標:", "type": "int", "default": str(step.params["y"])},
                    {"key": "click_type", "label": "クリックタイプ:", "type": "combobox", "default": step.params.get("click_type", "single"), "values": AppConfig.CLICK_TYPES, "required": True},
                    {"key": "comment", "label": "メモ:", "type": "text", "default": step.comment},
                ]
            elif step.type == "coord_drag":
                fields = [
                    {"key": "start_x", "label": "開始X座標:", "type": "int", "default": str(step.params["start_x"])},
                    {"key": "start_y", "label": "開始Y座標:", "type": "int", "default": str(step.params["start_y"])},
                    {"key": "end_x", "label": "終了X座標:", "type": "int", "default": str(step.params["end_x"])},
                    {"key": "end_y", "label": "終了Y座標:", "type": "int", "default": str(step.params["end_y"])},
                    {"key": "duration", "label": "ドラッグ時間(秒):", "type": "float", "default": str(step.params["duration"]), "min": 0.1},
                    {"key": "comment", "label": "メモ:", "type": "text", "default": step.comment},
                ]
            elif step.type == "image_relative_right_click":
                fields = [
                    {"key": "threshold", "label": "信頼度(0.5-1.0):", "type": "float", "default": str(step.params["threshold"]), "min": 0.5, "max": 1.0},
                    {"key": "click_type", "label": "クリックタイプ:", "type": "combobox", "default": step.params.get("click_type", "right"), "values": AppConfig.CLICK_TYPES, "required": True},
                    {"key": "offset_x", "label": "Xオフセット(-9999-9999):", "type": "int", "default": str(step.params["offset_x"]), "min": -9999},
                    {"key": "offset_y", "label": "Yオフセット(-9999-9999):", "type": "int", "default": str(step.params["offset_y"]), "min": -9999},
                    {"key": "retry", "label": "リトライ回数(0-10):", "type": "int", "default": str(step.params["retry"]), "min": 0, "max": 10},
                    {"key": "delay", "label": "リトライ間隔(秒, 0.1-10):", "type": "float", "default": str(step.params["delay"]), "min": 0.1, "max": 10.0},
                    {"key": "comment", "label": "メモ:", "type": "text", "default": step.comment},
                ]
            elif step.type == "sleep":
                current_wait_type = step.params.get("wait_type", "sleep")
                wait_type_display = "時刻指定" if current_wait_type == "scheduled" else "スリープ(秒数指定)"
                fields = [
                    {"key": "wait_type", "label": "待機タイプ:", "type": "combobox", "values": ["スリープ(秒数指定)", "時刻指定"], "default": wait_type_display, "on_change": True},
                    {"key": "seconds", "label": "待ち時間(秒, 0.1-100):", "type": "float", "default": str(step.params.get("seconds", 1.0)), "min": 0.1, "max": 100.0, "show_condition": {"field": "wait_type", "value": "スリープ(秒数指定)"}},
                    {"key": "scheduled_time", "label": "実行時刻(HH:MM:SS):", "type": "str", "default": step.params.get("scheduled_time", ""), "help": "日次実行時刻を指定（例：14:30:00）", "show_condition": {"field": "wait_type", "value": "時刻指定"}},
                    {"key": "comment", "label": "メモ:", "type": "text", "default": step.comment},
                ]
            elif step.type in ["copy", "paste", "key"]:
                fields = [
                    {"key": "key", "label": "キー:", "type": "combobox", "default": step.params["key"], "values": AppConfig.KEY_OPTIONS, "required": True},
                    {"key": "comment", "label": "メモ:", "type": "text", "default": step.comment},
                ]
            elif step.type == "custom_text":
                fields = [
                    {"key": "text", "label": "入力文字列:", "type": "text", "default": step.params["text"], "required": True},
                    {"key": "comment", "label": "メモ:", "type": "text", "default": step.comment},
                ]
            elif step.type == "cmd_command":
                # 完了を待つの現在値を文字列に変換
                wait_completion_value = "待つ" if step.params.get("wait_completion", True) else "待たない"
                fields = [
                    {"key": "command", "label": "コマンド:", "type": "text", "height": 12, "default": step.params["command"], "required": True, 
                     "help": "複数行のコマンドを記述できます"},
                    {"key": "timeout", "label": "タイムアウト(秒):", "type": "int", "default": str(step.params.get("timeout", 30)), "min": 1, "required": True},
                    {"key": "wait_completion", "label": "完了を待つ:", "type": "combobox", "values": ["待つ", "待たない"], "default": wait_completion_value},
                    {"key": "comment", "label": "メモ:", "type": "text", "default": step.comment},
                ]
            elif step.type == "repeat_start":
                fields = [
                    {"key": "count", "label": "繰り返し回数:", "type": "int", "default": str(step.params["count"]), "min": 1, "required": True},
                    {"key": "comment", "label": "メモ:", "type": "text", "default": step.comment},
                ]
            elif step.type == "repeat_end":
                fields = [
                    {"key": "comment", "label": "メモ:", "type": "text", "default": step.comment},
                ]

            # 全ての編集ダイアログサイズを統一
            dialog_width, dialog_height = 700, 800
            
            dialog = ModernDialog(self.root, f"{self.get_type_display(step)}編集", fields, width=dialog_width, height=dialog_height)
            try:
                result = dialog.get_result()
                if result:
                    if step.type == "image_click":
                        step.params.update({
                            "threshold": result["threshold"],
                            "click_type": result["click_type"],
                            "retry": result["retry"],
                            "delay": result["delay"],
                        })
                    elif step.type == "coord_click":
                        step.params.update({
                            "x": result["x"],
                            "y": result["y"],
                            "click_type": result["click_type"],
                        })
                    elif step.type == "coord_drag":
                        step.params.update({
                            "start_x": result["start_x"],
                            "start_y": result["start_y"],
                            "end_x": result["end_x"],
                            "end_y": result["end_y"],
                            "duration": result["duration"],
                        })
                    elif step.type == "image_relative_right_click":
                        step.params.update({
                            "threshold": result["threshold"],
                            "click_type": result["click_type"],
                            "offset_x": result["offset_x"],
                            "offset_y": result["offset_y"],
                            "retry": result["retry"],
                            "delay": result["delay"],
                        })
                    elif step.type == "sleep":
                        if result["wait_type"] == "時刻指定":
                            if not result["scheduled_time"].strip():
                                self.show_error_with_sound("エラー", "時刻指定を選択した場合は実行時刻を入力してください。")
                                return
                            # HH:MM:SS形式の検証
                            try:
                                time_parts = result["scheduled_time"].split(':')
                                if len(time_parts) != 3:
                                    raise ValueError("時刻の形式が正しくありません")
                                int(time_parts[0])  # 時
                                int(time_parts[1])  # 分  
                                int(time_parts[2])  # 秒
                                step.params["scheduled_time"] = result["scheduled_time"]
                                step.params["wait_type"] = "scheduled"
                                # 秒数パラメータを削除
                                if "seconds" in step.params:
                                    del step.params["seconds"]
                            except ValueError:
                                self.show_error_with_sound("エラー", "時刻の形式が正しくありません。HH:MM:SS形式で入力してください。（例：14:30:00）")
                                return
                        else:
                            # スリープ（秒数指定）
                            step.params["seconds"] = result["seconds"]
                            step.params["wait_type"] = "sleep"
                            # scheduled_timeパラメータを削除
                            if "scheduled_time" in step.params:
                                del step.params["scheduled_time"]
                    elif step.type in ["copy", "paste", "key"]:
                        step.params["key"] = result["key"]
                    elif step.type == "custom_text":
                        step.params["text"] = result["text"]
                    elif step.type == "cmd_command":
                        step.params["command"] = result["command"]
                        step.params["timeout"] = result["timeout"]
                        step.params["wait_completion"] = result["wait_completion"] == "待つ"  # 文字列からboolに変換
                    elif step.type == "repeat_start":
                        step.params["count"] = result["count"]
                    elif step.type == "repeat_end":
                        pass  # repeat_endは特別なパラメータなし

                    step.comment = result["comment"]
                    status = "✅" if step.enabled else "❌"
                    # 編集後は実際のコメント内容を表示（get_clean_commentは使わない）
                    display_comment = step.comment if step.comment.strip() else "-"
                    self.tree.item(selected_items[0], values=(status, index, self.get_type_display(step), self.get_params_display(step), display_comment))
                    
                    # ステップが編集されたら設定コンボボックスをクリア（編集中状態を示す）
                    if hasattr(self, 'config_combo'):
                        self.config_combo.set("")
                    
                    logger.info(f"ステップ編集: 行番号={index}, step={step}")
            except Exception as e:
                logger.error(f"ステップ編集エラー: 行番号={index}, error={e}")
                self.show_error_with_sound("エラー", f"ステップの編集に失敗しました: 行番号={index}, エラー: {e}")
                # Select the erroneous step in the Treeview
                try:
                    self.tree.selection_set(selected_items[0])
                    self.tree.see(selected_items[0])
                    logger.info(f"エラー行を選択: 行番号={index}")
                except Exception as select_error:
                    logger.error(f"エラー行の選択エラー: 行番号={index}, error={str(select_error)}")
        except Exception as e:
            logger.error(f"ステップ編集エラー: {e}")
            self.show_error_with_sound("エラー", f"ステップの編集に失敗しました: 行番号なし, エラー: {e}")






    def _validate_selection(self):
        """選択された項目を確認"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "編集するステップを選択してください")
            return None
        return selected_items

    def move_up(self, event=None):
        """ステップを上に移動"""
        try:
            selected_items = self._validate_selection()
            if selected_items:
                index = self.tree.index(selected_items[0])
                if index > 0:
                    self.steps[index], self.steps[index - 1] = self.steps[index - 1], self.steps[index]
                    self.refresh_tree()
                    self.tree.selection_set(self.tree.get_children()[index - 1])
                    logger.info(f"ステップを上に移動: index={index}")
        except Exception as e:
            logger.error(f"ステップ移動エラー: {e}")
            self.show_error_with_sound("エラー", f"ステップの移動に失敗しました: {e}")

    def move_down(self, event=None):
        """ステップを下に"""
        try:
            selected_items = self._validate_selection()
            if selected_items:
                index = self.tree.index(selected_items[0])
                if index < len(self.steps) - 1:
                    self.steps[index], self.steps[index + 1] = self.steps[index + 1], self.steps[index]
                    self.refresh_tree()
                    self.tree.selection_set(self.tree.get_children()[index + 1])
                    logger.info(f"ステップを下に移動: index={index}")
        except Exception as e:
            logger.error(f"ステップ下移動エラー: {e}")
            self.show_error_with_sound("エラー", f"ステップの移動に失敗しました: {e}")


    def refresh_tree(self):
        """ツリーをリフレッシュ"""
        try:
            # 既存のアイテムを削除
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # full_valuesをクリア
            self.drag_drop_tree.full_values.clear()
            
            for index, step in enumerate(self.steps, start=1):  # 1-based index
                status = "✅" if step.enabled else "❌"
                
                # 完全なテキストを取得
                params_full = self.get_params_display(step)
                comment_full = step.comment
                
                # 省略表示を適用
                params_display = self.drag_drop_tree.elide_to_fit(params_full, "Params")
                comment_display = self.drag_drop_tree.elide_to_fit(comment_full, "Comment")
                
                # アイテムを挿入
                item_id = self.tree.insert("", tk.END, values=(status, index, self.get_type_display(step), params_display, comment_display))
                
                # 完全なテキストを保存
                self.drag_drop_tree.full_values[item_id] = {
                    'Params': params_full,
                    'Comment': comment_full
                }
                
                # 無効行のタグを適用
                self.drag_drop_tree.apply_row_state_tags(item_id, step.enabled)
                
            logger.info("ツリーを更新しました")
            # ツリー更新のアニメーション効果
            self.animate_tree_update()
        except Exception as e:
            logger.error(f"ツリー更新エラー: {e}")
            self.show_error_with_sound("エラー", f"ツリーの更新に失敗しました: 行番号なし, エラー: {e}")

    def highlight_current_step(self, step_index: int):
        """実行中のステップをハイライト表示"""
        try:
            # 既存の選択をクリア
            for item in self.tree.selection():
                self.tree.selection_remove(item)
            
            # 指定されたステップを選択・フォーカス
            children = self.tree.get_children()
            if 0 <= step_index < len(children):
                item_id = children[step_index]
                self.tree.selection_set(item_id)
                self.tree.focus(item_id)
                self.tree.see(item_id)  # スクロールして表示
                
                # ハイライト用のタグを設定（安全な方法）
                try:
                    self.tree.item(item_id, tags=("current_step",))
                except Exception as tag_error:
                    logger.debug(f"ステップタグ設定エラー（無視可能）: {tag_error}")
                
                # 画像プレビューを同期
                self.update_image_preview(step_index)
                
        except Exception as e:
            logger.error(f"ステップハイライトエラー: {e}")

    def update_image_preview(self, step_index: int):
        """指定されたステップの画像プレビューを更新"""
        try:
            if hasattr(self, 'image_preview') and 0 <= step_index < len(self.steps):
                step = self.steps[step_index]
                image_path = step.get_preview_image_path()
                if image_path:
                    self.image_preview.show_image(image_path)
                    logger.debug(f"画像プレビュー更新: ステップ{step_index+1}, path={image_path}")
                else:
                    logger.debug(f"画像プレビュー更新スキップ: ステップ{step_index+1}は画像ステップではありません")
        except Exception as e:
            logger.error(f"画像プレビュー更新エラー: step_index={step_index}, error={e}")

    def clear_current_step_highlight(self):
        """現在のステップハイライトをクリア"""
        try:
            # 全ての選択をクリア
            for item in self.tree.selection():
                self.tree.selection_remove(item)
            
            # 実行マーカーをクリア（安全な方法）
            for item in self.tree.get_children():
                try:
                    self.tree.item(item, tags=())
                except Exception as tag_error:
                    logger.debug(f"ステップタグクリアエラー（無視可能）: {tag_error}")
                    
        except Exception as e:
            logger.error(f"ステップハイライトクリアエラー: {e}")

    def take_screenshot_and_save(self):
        """スクリーンショットを撮影してファイルに保存"""
        try:
            logger.info("スクリーンショット処理開始")
            
            # PyAutoGUIを使って簡単にスクリーンショットを撮影
            import pyautogui
            from datetime import datetime
            
            logger.info("PyAutoGUIでスクリーンショット撮影開始")
            screenshot = pyautogui.screenshot()
            logger.info("スクリーンショット撮影完了")
            
            # スクリーンショット保存ディレクトリを作成
            import os
            screenshot_dir = os.path.join(os.getcwd(), "screenshots")
            if not os.path.exists(screenshot_dir):
                os.makedirs(screenshot_dir)
                logger.info("スクリーンショットディレクトリを作成: " + screenshot_dir)
            
            # ファイル名生成
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = "screenshot_" + timestamp + ".png"
            file_path = os.path.join(screenshot_dir, filename)
            logger.info("保存先ファイルパス: " + file_path)
            
            # スクリーンショットを保存
            logger.info("ファイル保存開始")
            screenshot.save(file_path, "PNG")
            logger.info("ファイル保存完了")
            logger.info("スクリーンショット保存成功: " + file_path)
                
        except Exception as main_error:
            logger.error("スクリーンショット処理でエラーが発生: " + str(type(main_error).__name__))
            try:
                logger.error("エラー詳細: " + str(main_error))
            except:
                logger.error("エラー詳細の取得に失敗")
    
    def add_specific_key_step(self, key: str, comment: str):
        """特定のキーステップを追加する統合メソッド"""
        try:
            # スクリーンショットの場合は特別な待機コメントを設定
            if key.lower() == "ctrl+shift+f12":
                sleep_comment = "スクリーンショット撮影後待機"
            else:
                sleep_comment = f"{comment}後待機"
            
            self.add_step(Step("key", params={"key": key}, comment=comment))
            self.add_step(Step("sleep", params={"seconds": 0.5}, comment=sleep_comment))
            logger.info(f"キーステップを追加しました: {comment}")
        except Exception as e:
            logger.error(f"キーステップ追加エラー ({comment}): " + str(e))

    def add_step_screenshot(self):
        """スクリーンショット撮影ステップを追加（後方互換性のため残存）"""
        self.add_specific_key_step("ctrl+shift+f12", "スクリーンショット撮影")
    
    def test_key_binding(self):
        """キーバインドのテスト用メソッド"""
        messagebox.showinfo("テスト", "キーバインドが正常に動作しています！")
        logger.info("キーバインドテスト実行: Ctrl+Shift+T")
    
    def save_config(self):
        """設定をファイルに保存（自動読み込み対応改善版）"""
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".json", 
                filetypes=[("JSON files", "*.json")],
                title="設定ファイルを保存"
            )
            if file_path:
                # 設定データを作成
                config = {
                    "steps": [step.__dict__ for step in self.steps],
                    "last_monitor": self.selected_monitor,
                    "loop_count": self.loop_count,
                    "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "step_count": len(self.steps)
                }
                
                # 設定ファイルを保存
                with open(file_path, "w", encoding="utf-8") as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                
                # 最後に使用したファイル情報を記録（次回自動読み込み用）
                last_config = {
                    "last_file": file_path, 
                    "last_monitor": self.selected_monitor, 
                    "loop_count": self.loop_count,
                    "last_saved": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "file_name": os.path.basename(file_path)
                }
                
                with open(self.config_file, "w", encoding="utf-8") as f:
                    json.dump(last_config, f, ensure_ascii=False, indent=2)
                
                # 成功メッセージ
                success_msg = f"設定保存完了: {os.path.basename(file_path)} ({len(self.steps)}ステップ)"
                logger.info(success_msg)
                self.update_status(success_msg)
                
                # 設定ファイル一覧を更新
                self.update_config_list()
                
                # 音声フィードバック（成功）
                try:
                    import winsound
                    winsound.MessageBeep(winsound.MB_ICONASTERISK)  # 成功音
                except:
                    pass
                    
            else:
                logger.info("設定保存がキャンセルされました")
                self.update_status("設定保存がキャンセルされました")
                
        except Exception as e:
            logger.error(f"設定保存エラー: {e}")
            self.show_error_with_sound("エラー", f"設定の保存に失敗しました: {e}")

    def _validate_config_structure(self, config: dict) -> bool:
        """設定ファイルの構造を検証"""
        try:
            # 基本構造チェック
            if not isinstance(config, dict):
                return False
            
            # 必須フィールドの存在チェック
            if "steps" not in config:
                return False
            
            steps = config["steps"]
            if not isinstance(steps, list):
                return False
            
            # 各ステップの基本構造をチェック
            for i, step in enumerate(steps):
                if not isinstance(step, dict):
                    logger.warning(f"ステップ {i}: 辞書形式ではありません")
                    return False
                
                # 必須フィールド
                if "type" not in step or "params" not in step:
                    logger.warning(f"ステップ {i}: 必須フィールド (type/params) がありません")
                    return False
                
                if not isinstance(step["params"], dict):
                    logger.warning(f"ステップ {i}: paramsが辞書形式ではありません")
                    return False
                
                # ステップタイプの検証（実装済みのすべてのタイプを含める）
                valid_types = [
                    "image_click", "coord_click", "coord_drag", "image_right_click", "image_relative_right_click", 
                    "sleep", "custom_text", "cmd", "cmd_command", "key_action", "key",
                    "copy", "paste", "repeat_start", "repeat_end"
                ]
                if step["type"] not in valid_types:
                    logger.warning(f"ステップ {i}: 不正なステップタイプ '{step['type']}'")
                    return False
            
            # 設定値の範囲チェック
            if "last_monitor" in config:
                monitor_val = config["last_monitor"]
                if not isinstance(monitor_val, int) or monitor_val < 0 or monitor_val > 10:
                    logger.warning(f"不正なモニター値: {monitor_val}")
                    config["last_monitor"] = 0  # デフォルト値に修正
            
            if "loop_count" in config:
                loop_val = config["loop_count"]
                if not isinstance(loop_val, int) or loop_val < 1 or loop_val > 10000:
                    logger.warning(f"不正なループ回数: {loop_val}")
                    config["loop_count"] = 1  # デフォルト値に修正
            
            return True
            
        except Exception as e:
            logger.error(f"設定ファイル検証エラー: {e}")
            return False

    def load_config(self):
        """設定をファイルから読み込む（自動読み込み対応改善版）"""
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("JSON files", "*.json")],
                title="設定ファイルを読み込み"
            )
            if file_path:
                with open(file_path, "r", encoding="utf-8") as f:
                    config = json.load(f)
                
                # 設定ファイル構造の検証
                if not self._validate_config_structure(config):
                    error_msg = "設定ファイルの構造が正しくありません"
                    logger.error(f"{error_msg}: {file_path}")
                    messagebox.showerror("読み込みエラー", error_msg)
                    return
                    
                # ステップデータの読み込みと補正
                steps = []
                for step_data in config.get("steps", []):
                    # 旧い形式のimage_clickステップを補正
                    if step_data.get("type") == "image_click" and "click_type" not in step_data.get("params", {}):
                        step_data["params"]["click_type"] = "single"
                        logger.info(f"旧い形式のimage_clickステップを補正: click_type='single'を追加")
                    steps.append(Step(**step_data))
                
                self.steps = steps
                self.selected_monitor = int(config.get("last_monitor", 0))
                self.loop_count = int(config.get("loop_count", 1))
                
                # UIを更新
                if hasattr(self, 'monitor_var'):
                    self.monitor_var.set(str(self.selected_monitor))
                self.select_monitor(str(self.selected_monitor))
                self.refresh_tree()
                
                # 次回自動読み込み用に記録
                last_config = {
                    "last_file": file_path,
                    "last_monitor": self.selected_monitor,
                    "loop_count": self.loop_count,
                    "last_loaded": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "file_name": os.path.basename(file_path)
                }
                
                with open(self.config_file, "w", encoding="utf-8") as f:
                    json.dump(last_config, f, ensure_ascii=False, indent=2)
                
                # 成功メッセージ
                success_msg = f"設定読み込み完了: {os.path.basename(file_path)} ({len(steps)}ステップ)"
                logger.info(success_msg)
                self.update_status(success_msg)
                
                # 設定ファイル一覧を更新
                self.update_config_list()
                
                # 音声フィードバック（成功）
                try:
                    import winsound
                    winsound.MessageBeep(winsound.MB_ICONASTERISK)
                except:
                    pass
                    
            else:
                logger.info("設定読み込みがキャンセルされました")
                self.update_status("設定読み込みがキャンセルされました")
                
        except Exception as e:
            logger.error(f"設定読み込みエラー: {e}")
            self.show_error_with_sound("エラー", f"設定の読み込みに失敗しました: {e}")

    def update_config_list(self):
        """設定ファイル一覧を更新"""
        try:
            # config_comboが初期化されているかチェック
            if not hasattr(self, 'config_combo'):
                logger.warning("config_combo が初期化されていません")
                return
                
            config_dir = os.path.join(os.path.dirname(__file__), "config")
            if not os.path.exists(config_dir):
                logger.warning(f"configディレクトリが存在しません: {config_dir}")
                return
            
            # config配下のJSONファイル一覧を取得
            config_files = []
            for file in os.listdir(config_dir):
                if file.endswith(".json") and file != "last_config.json":
                    config_files.append(file)
            
            # ファイル名順でソート
            config_files.sort()
            logger.info(f"設定ファイル一覧: {config_files}")
            
            # コンボボックスを更新
            self.config_combo['values'] = config_files
            
            # 現在のファイルを選択状態にする
            try:
                if os.path.exists(self.config_file):
                    with open(self.config_file, "r", encoding="utf-8") as f:
                        last_config = json.load(f)
                    current_file = last_config.get("file_name", "")
                    if current_file in config_files:
                        self.config_combo.set(current_file)
                        logger.info(f"現在のファイルを選択: {current_file}")
                    elif config_files:
                        self.config_combo.set(config_files[0])
                        logger.info(f"デフォルトファイルを選択: {config_files[0]}")
                elif config_files:
                    self.config_combo.set(config_files[0])
                    logger.info(f"デフォルトファイルを選択: {config_files[0]}")
            except Exception as inner_e:
                logger.error(f"現在ファイル選択エラー: {inner_e}")
                if config_files:
                    self.config_combo.set(config_files[0])
                    
        except Exception as e:
            logger.error(f"設定ファイル一覧更新エラー: {e}")
            import traceback
            traceback.print_exc()
    
    def on_config_selected(self, event=None):
        """設定ファイルが選択された時の処理"""
        selected_file = self.config_combo.get()
        if selected_file:
            config_dir = os.path.join(os.path.dirname(__file__), "config")
            file_path = os.path.join(config_dir, selected_file)
            if os.path.exists(file_path):
                self.load_config_file(file_path)
    
    def on_theme_selected(self, event=None):
        """テーマが選択された時の処理"""
        selected_theme_name = self.theme_combo.get()
        # テーマ名からキーを見つける
        for key, theme in AppConfig.THEMES.items():
            if theme['name'] == selected_theme_name:
                if AppConfig.set_theme(key):
                    # スムーズなテーマ切り替えアニメーション
                    self.animate_theme_transition()
                    # テーマ適用
                    AppConfig.apply_dark_theme(self.root)
                    # ステータス更新
                    self.update_status(f"テーマを変更しました: {selected_theme_name}")
                    logger.info(f"テーマ変更: {selected_theme_name}")
                break

    def animate_theme_transition(self):
        """テーマ切り替えのスムーズアニメーション"""
        try:
            # 短時間のフェード効果で視覚的フィードバック
            self.root.configure(relief='sunken')
            self.root.after(100, lambda: self.root.configure(relief='flat'))
            
            # ステータスバーにアニメーション効果
            if hasattr(self, 'status_label'):
                original_bg = self.status_label.cget('background')
                self.status_label.configure(background=AppConfig.THEME['bg_accent'])
                self.root.after(200, lambda: self.status_label.configure(background=original_bg))
        except Exception as e:
            logger.debug(f"テーマ切り替えアニメーションエラー: {e}")
    
    
    
    
    def load_config_file(self, file_path):
        """指定されたファイルパスの設定を読み込み"""
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                config = json.load(f)
            
            # 設定ファイル構造の検証
            if not self._validate_config_structure(config):
                error_msg = "設定ファイルの構造が正しくありません"
                logger.error(f"{error_msg}: {file_path}")
                messagebox.showerror("読み込みエラー", error_msg)
                return
                
            # ステップデータの読み込みと補正
            steps = []
            for step_data in config.get("steps", []):
                # 旧い形式のimage_clickステップを補正
                if step_data.get("type") == "image_click" and "click_type" not in step_data.get("params", {}):
                    step_data["params"]["click_type"] = "single"
                    logger.info(f"旧い形式のimage_clickステップを補正: click_type='single'を追加")
                steps.append(Step(**step_data))
            
            self.steps = steps
            self.selected_monitor = config.get("last_monitor", 0)
            self.loop_count = int(config.get("loop_count", 1))
            
            # UIを更新
            if hasattr(self, 'monitor_var'):
                self.monitor_var.set(str(self.selected_monitor))
            self.select_monitor(str(self.selected_monitor))
            self.refresh_tree()
            
            # 次回自動読み込み用に記録
            last_config = {
                "last_file": file_path,
                "last_monitor": self.selected_monitor,
                "loop_count": self.loop_count,
                "last_loaded": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "file_name": os.path.basename(file_path)
            }
            
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(last_config, f, ensure_ascii=False, indent=2)
            
            # 成功メッセージ
            success_msg = f"設定読み込み完了: {os.path.basename(file_path)} ({len(steps)}ステップ)"
            logger.info(success_msg)
            self.update_status(success_msg)
            
            # コンボボックスの選択状態を更新
            self.config_combo.set(os.path.basename(file_path))
            
            # 音声フィードバック（成功）
            try:
                import winsound
                winsound.MessageBeep(winsound.MB_OK)
            except:
                pass
                
        except Exception as e:
            logger.error(f"設定ファイル読み込みエラー: {file_path}, エラー: {e}")
            self.show_error_with_sound("エラー", f"設定ファイルの読み込みに失敗しました: {e}")

    def load_last_config(self):
        """最後の設定を自動読み込み（改善版）"""
        try:
            if os.path.exists(self.config_file):
                try:
                    with open(self.config_file, "r", encoding="utf-8") as f:
                        config = json.load(f)
                    
                    # 基本設定を復元
                    last_file = config.get("last_file")
                    self.selected_monitor = int(config.get("last_monitor", 0))
                    self.loop_count = int(config.get("loop_count", 1))
                    
                    # モニター設定をUIに反映
                    if hasattr(self, 'monitor_var'):
                        self.monitor_var.set(str(self.selected_monitor))
                    
                    # 前回のファイルが存在する場合は自動読み込み
                    if last_file and os.path.exists(last_file):
                        try:
                            with open(last_file, "r", encoding="utf-8") as f:
                                data = json.load(f)
                            
                            # ステップデータの読み込みと補正
                            steps = []
                            for step_data in data.get("steps", []):
                                # 旧形式のimage_clickステップを補正
                                if step_data.get("type") == "image_click" and "click_type" not in step_data.get("params", {}):
                                    step_data["params"]["click_type"] = "single"
                                    logger.info(f"旧形式のimage_clickステップを補正: click_type='single'を追加")
                                steps.append(Step(**step_data))
                            
                            self.steps = steps
                            self.loop_count = int(data.get("loop_count", self.loop_count))
                            
                            # UIを更新
                            self.refresh_tree()
                            
                            # 成功メッセージ
                            status_msg = f"前回の設定を自動読み込み完了: {os.path.basename(last_file)} ({len(steps)}ステップ)"
                            logger.info(status_msg)
                            
                            # ステータスバーに表示（遅延実行）
                            self.root.after(500, lambda: self.update_status(status_msg))
                            
                        except Exception as e:
                            logger.error(f"前回ファイルの読み込みエラー: {e}")
                            self.update_status(f"前回ファイルの読み込みに失敗: {os.path.basename(last_file)}")
                            
                    elif last_file:
                        logger.info(f"前回の設定ファイルが見つかりません: {last_file}")
                        self.update_status("前回の設定ファイルが見つかりません")
                    else:
                        logger.info("前回保存された設定ファイルはありません")
                        self.update_status("新規セッション開始")
                    
                    # モニター選択を確実に実行
                    self.select_monitor(str(self.selected_monitor))
                    
                except json.JSONDecodeError as e:
                    logger.error(f"設定ファイルの形式エラー: {e}")
                    self.update_status("設定ファイルの形式が正しくありません")
                    self.select_monitor("0")
                    
            else:
                logger.info("初回起動: 設定ファイルが存在しません")
                self.update_status("初回起動 - 新規セッション開始")
                self.selected_monitor = 0
                self.loop_count = 1
                if hasattr(self, 'monitor_var'):
                    self.monitor_var.set("0")
                self.select_monitor("0")
                
        except Exception as e:
            logger.error(f"設定読み込み処理エラー: {e}")
            self.update_status("設定の読み込みに問題が発生しました")
            # フォールバック処理
            self.selected_monitor = 0
            self.loop_count = 1
            if hasattr(self, 'monitor_var'):
                self.monitor_var.set("0")
            self.select_monitor("0")

    def run_all_steps(self):
        """すべてのステップを実行"""
        try:
            self.loop_count = 1  # 繰り返しアクションを使用するため固定
            if self.loop_count < 0:
                raise ValueError("ループ回数は0以上の整数で設定してください")
            self.running = True
            
            # ボタン状態を更新
            self.update_execution_buttons(True)
            
            # 進捗追跡を開始（繰り返し回数を考慮した総ステップ数を自動計算）
            self.start_execution_tracking()
            
            # 実行開始時にハイライトをクリアして初期化
            self.clear_current_step_highlight()
            status_text = "▶️ 実行中..."
            if hasattr(self, 'main_status_label'):
                self.main_status_label.config(text=status_text)
            elif hasattr(self, 'status_label'):
                self.status_label.config(text=status_text)
            threading.Thread(target=self._run_all_steps, daemon=True).start()
            logger.info(f"実行開始: ループ回数={self.loop_count}")
        except ValueError as e:
            logger.error(f"実行開始エラー: {e}")
            self.show_error_with_sound("エラー", str(e))

    def run_from_selected(self, event=None):
        """選択したステップから実行"""
        try:
            # 選択されたアイテムを取得
            selected = self.tree.selection()
            if not selected:
                messagebox.showinfo("情報", "実行開始するステップを選択してください")
                return
            
            # 選択されたステップのインデックスを取得
            start_index = self.tree.index(selected[0])
            
            # 通常の選択からの実行
            logger.info(f"選択実行: ステップ{start_index+1}から")
            self.execution_start_index = start_index
                
            self.loop_count = 1  # 選択から実行は1回のみ
            self.running = True
            
            if hasattr(self, 'main_run_btn'):
                self.main_run_btn.configure(text="⏸️ 実行中", state="disabled")
            elif hasattr(self, 'start_button'):
                self.start_button.configure(text="実行中", state="disabled")
            if hasattr(self, 'stop_button'):
                self.stop_button.configure(state="normal")
            
            status_text = f"▶️ ステップ{start_index+1}から実行中..."
            if hasattr(self, 'main_status_label'):
                self.main_status_label.config(text=status_text)
            elif hasattr(self, 'status_label'):
                self.status_label.config(text=status_text)
                
            # 選択されたステップから実行開始
            threading.Thread(target=self._run_from_index, daemon=True).start()
            logger.info(f"選択ステップから実行開始: 開始インデックス={start_index}")
            
        except Exception as e:
            logger.error(f"選択ステップ実行エラー: {e}")
            self.show_error_with_sound("エラー", f"選択ステップからの実行に失敗しました: {e}")
    
    def _run_from_index(self):
        """指定インデックスから実行"""
        try:
            self._execute_loop(self._run_steps_from_index, "選択ステップ")
            logger.info("選択ステップからの実行処理を開始しました")
        except Exception as e:
            logger.error(f"選択ステップ実行エラー: {e}")
            self.update_status("❌ エラー発生")
            self.show_error_with_sound("エラー", f"選択ステップからの実行に失敗しました: {e}")
    
    def _run_steps_from_index(self):
        """指定インデックスからステップ実行"""
        try:
            monitor_index = self.selected_monitor
            return self._execute_steps_for_monitor_from_index(monitor_index, self.execution_start_index)
        except Exception as e:
            logger.error(f"選択インデックス実行エラー: {e}")
            return False
    
    def _execute_steps_for_monitor_from_index(self, monitor_index: int, start_index: int) -> bool:
        """モニターごとのステップ実行（指定インデックスから）"""
        try:
            # monitor_indexが文字列の場合は整数に変換
            if isinstance(monitor_index, str):
                monitor_index = int(monitor_index)
            # 指定インデックス以降のステップで実行計画生成
            steps_from_index = self.steps[start_index:]
            execution_plan = self._generate_execution_plan_from_steps(steps_from_index, start_index)
            
            # 有効ステップの総数を計算
            total_valid_steps = sum(1 for step_index, _ in execution_plan 
                                   if self.steps[step_index].enabled)
            
            # 実行済みステップのカウンター
            executed_steps = 0
            
            for exec_index, (step_index, repeat_iter) in enumerate(execution_plan, start=1):
                if not self.running:
                    self.update_status("⛔ 実行中断")
                    logger.info("ステップ実行が中断されました")
                    return False
                
                step = self.steps[step_index]
                
                # 無効なステップをスキップ
                if not step.enabled:
                    logger.info(f"ステップスキップ: 行番号={step_index+1}, type={step.type} (無効)")
                    continue
                
                # 繰り返し表示
                repeat_text = f" (繰り返し{repeat_iter+1}回目)" if repeat_iter > 0 else ""
                self.update_status(f"▶️ ステップ {step_index+1}/{len(self.steps)}: {self.get_type_display(step)}{repeat_text}")
                logger.info(f"ステップ実行開始: 行番号={step_index+1}, type={step.type}, repeat_iter={repeat_iter}")
                
                # 有効ステップの場合、進捗を更新
                if step.enabled:
                    executed_steps += 1
                    # 進行状況を更新
                    self.progress_var.set(f"{executed_steps}/{total_valid_steps}")
                
                # 実行中のステップをハイライト表示
                self.highlight_current_step(step_index)
                
                # プログレス可視化を更新
                current_progress = executed_steps / total_valid_steps if total_valid_steps > 0 else 0
                step_name = f"{step.comment}" if step.comment else f"{step.type.upper()}"
                self.update_realtime_info(step_name, current_progress)
                self.update_execution_stats(step_name)
                
                try:
                    if step.type == "repeat_start":
                        logger.info(f"繰り返し開始: {step.params['count']}回")
                    elif step.type == "repeat_end":
                        logger.info(f"繰り返し終了")
                    elif step.type == "image_click":
                        self._execute_image_click(step, monitor_index)
                    elif step.type == "coord_click":
                        self._execute_coord_click(step)
                    elif step.type == "coord_drag":
                        self._execute_coord_drag(step)
                    elif step.type == "image_relative_right_click":
                        self._execute_image_right_click(step, monitor_index)
                    elif step.type == "sleep":
                        self._execute_sleep(step)
                    elif step.type == "custom_text":
                        self._execute_custom_text(step)
                    elif step.type == "cmd_command":
                        self._execute_cmd_command(step)
                    else:
                        self._execute_key_action(step)
                    logger.info(f"ステップ実行成功: 行番号={step_index+1}, step={step}")
                    
                    # 成功時の統計とアニメーション更新
                    self.update_execution_stats(success=True)
                    self.animate_step_completion(step_index, success=True)
                except Exception as e:
                    self.update_status(f"❌ エラー発生: 行番号 {step_index+1}")
                    logger.error(f"ステップ実行エラー: 行番号={step_index+1}, step={step}, error={str(e)}")
                    
                    # スマートエラー分析とダイアログ表示
                    analysis = self.analyze_error(e, step, step_index+1)
                    self.show_error_dialog(analysis, step, step_index+1)
                    
                    
                    self.running = False
                    self.update_execution_buttons(False)
                    # Select the erroneous step in the Treeview
                    try:
                        children = self.tree.get_children()
                        if step_index < len(children):
                            self.tree.selection_set(children[step_index])
                            self.tree.see(children[step_index])
                            self.update_image_preview(step_index)  # 画像プレビューも同期
                            logger.info(f"エラー行を選択: 行番号={step_index+1}")
                        else:
                            logger.warning(f"エラー行の選択に失敗: 行番号={step_index+1}, ツリーアイテム数={len(children)}")
                    except Exception as select_error:
                        logger.error(f"エラー行の選択エラー: 行番号={step_index+1}, error={str(select_error)}")
                    # エラー統計を更新
                    self.update_execution_stats(error=True)
                    self.animate_step_completion(step_index, success=False)
                    
                    # エラー時にハイライトをクリア（コメントアウト - エラー行を視認しやすくするため）
                    # self.clear_current_step_highlight()
                    return False
            # 実行完了時にハイライトをクリア
            self.clear_current_step_highlight()
            
            # 完了時の進捗表示を更新（100%完了）
            self.progress_var.set(f"{total_valid_steps}/{total_valid_steps}")
            self.update_realtime_info("実行完了", 1.0)
            if hasattr(self, 'main_progress_bar') and self.main_progress_bar:
                self.update_progress_bar(self.main_progress_bar, 1.0, "実行完了 ✅", animate=True)
            
            logger.info(f"モニター[{monitor_index}]の選択ステップ実行完了")
            return True
        except Exception as e:
            self.update_status("❌ エラー発生")
            logger.error(f"モニター[{monitor_index}]選択ステップ実行エラー: {e}")
            self.show_error_with_sound("エラー", f"モニター[{monitor_index}]の選択ステップ実行に失敗しました: 行番号なし, エラー: {e}")
            self.running = False
            self.update_execution_buttons(False)
            # エラー時にハイライトをクリア（コメントアウト - エラー行を視認しやすくするため）
            # self.clear_current_step_highlight()
            return False
    
    def _generate_execution_plan_from_steps(self, steps, offset=0):
        """指定されたステップリストから実行計画生成（ネスト対応）"""
        return self._expand_nested_loops_from_steps(steps, 0, len(steps), offset)

    def _expand_nested_loops_from_steps(self, steps, start_idx, end_idx, offset=0, nest_level=0):
        """指定されたステップリストからネストしたループを再帰的に展開"""
        execution_plan = []
        i = start_idx
        
        while i < end_idx:
            step = steps[i]
            actual_index = i + offset  # 実際のステップインデックス
            
            if step.type == "repeat_start":
                repeat_count = step.params.get('count', 1)
                execution_plan.append((actual_index, nest_level))  # repeat_start自体を追加
                
                # 対応するrepeat_endを見つける
                nest_depth = 1
                end_pos = i + 1
                while end_pos < end_idx and nest_depth > 0:
                    if steps[end_pos].type == "repeat_start":
                        nest_depth += 1
                    elif steps[end_pos].type == "repeat_end":
                        nest_depth -= 1
                    end_pos += 1
                
                if nest_depth == 0:  # 対応するrepeat_endが見つかった
                    end_pos -= 1  # repeat_endのインデックスに調整
                    
                    # 繰り返し処理
                    for repeat_iter in range(repeat_count):
                        # ループ内容を再帰的に展開
                        inner_plan = self._expand_nested_loops_from_steps(
                            steps, i + 1, end_pos, offset, nest_level + repeat_iter + 1
                        )
                        execution_plan.extend(inner_plan)
                        # 各繰り返しの最後にrepeat_endも実行
                        execution_plan.append((end_pos + offset, nest_level + repeat_iter + 1))
                    
                    i = end_pos + 1  # repeat_endの次に進む
                else:
                    raise ValueError(f"対応するrepeat_endが見つかりません: ステップ{actual_index}")
                    
            elif step.type == "repeat_end":
                # 単体のrepeat_endは無視（親の処理で対応済み）
                i += 1
            else:
                # 通常のステップ
                execution_plan.append((actual_index, nest_level))
                i += 1
        
        return execution_plan

    def run_all_monitors(self, event=None):
        """全モニターで検索を実行"""
        try:
            self.loop_count = 1  # 繰り返しアクションを使用するため固定
            if self.loop_count < 0:
                raise ValueError("ループ回数は0以上の整数で設定してください")
            self.running = True
            if hasattr(self, 'main_run_btn'):
                self.main_run_btn.configure(text="⏸️ 実行中", state="disabled")
            elif hasattr(self, 'start_button'):
                self.start_button.configure(text="実行中", state="disabled")
            if hasattr(self, 'stop_button'):
                self.stop_button.configure(state="normal")
            self.update_status("🔍 全モニター検索中...")
            threading.Thread(target=self._run_steps_all_monitors, daemon=True).start()
            logger.info(f"全モニター検索開始: ループ回数={self.loop_count}")
        except ValueError as e:
            logger.error(f"全モニター検索開始エラー: {e}")
            self.show_error_with_sound("エラー", str(e))

    def emergency_stop(self):
        """ESC緊急停止"""
        if self.running:
            logger.warning("ESC緊急停止が実行されました")
            self.stop_execution()
            # 緊急停止の視覚的フィードバック
            if hasattr(self, 'main_status_label'):
                self.main_status_label.config(text="🚨 ESC緊急停止しました", fg='#e74c3c')
                self.root.after(3000, lambda: self.main_status_label.config(fg='#ffffff'))
            # システム音で停止を知らせる
            try:
                import os
                os.system("echo \a")
            except:
                pass

    def stop_execution(self, event=None):
        """実行を停止"""
        try:
            self.running = False
            
            # ボタン状態を更新
            self.update_execution_buttons(False)
            
            # 実行停止時にハイライトをクリア
            self.clear_current_step_highlight()
            status_text = "⏹️ 停止しました"
            if hasattr(self, 'main_status_label'):
                self.main_status_label.config(text=status_text)
            elif hasattr(self, 'status_label'):
                self.status_label.config(text=status_text)
            logger.info("実行を停止しました")
        except Exception as e:
            logger.error(f"実行停止エラー: {e}")
            self.show_error_with_sound("エラー", f"実行の停止に失敗しました: {e}")

    def get_monitor_region(self, monitor_index: int) -> Tuple[int, int, int, int]:
        """モニター領域を取得"""
        try:
            # monitor_indexが文字列の場合は整数に変換
            if isinstance(monitor_index, str):
                monitor_index = int(monitor_index)
            
            monitor = self.monitors[monitor_index]
            scale = self.dpi_scale / 100.0
            region = (
                int(monitor.x // scale),
                int(monitor.y // scale),
                int(monitor.width // scale),
                int(monitor.height // scale),
            )
            logger.info(f"モニター領域取得: index={monitor_index}, region={region}")
            return region
        except Exception as e:
            logger.error(f"モニター領域取得エラー: index={monitor_index}, error={e}")
            raise ValueError(f"モニター領域の取得に失敗しました: {e}")

    @contextmanager
    def mss_context(self):
        """mssインスタンスをスレッドセーフに管理"""
        sct = mss.mss()
        try:
            yield sct
        except Exception as e:
            logger.error(f"mssコンテキストエラー: {e}")
            raise
        finally:
            sct.close()

    def capture_screenshot(self, monitor_index: int) -> np.ndarray:
        """スクリーンショットを撮影"""
        try:
            # monitor_indexが文字列の場合は整数に変換
            if isinstance(monitor_index, str):
                monitor_index = int(monitor_index)
                
            with self.mss_context() as sct:
                monitor = sct.monitors[monitor_index + 1]
                screenshot = sct.grab({
                    "left": monitor["left"],
                    "top": monitor["top"],
                    "width": monitor["width"],
                    "height": monitor["height"],
                })
                img = Image.frombytes("RGB", screenshot.size, screenshot.rgb)
                img_np = np.array(img)
                img_np = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)
                logger.info(f"スクリーンショット取得成功: モニター[{monitor_index}]")
                return img_np
        except Exception as e:
            logger.error(f"スクリーンショットエラー: モニター[{monitor_index}], error={str(e)}")
            raise RuntimeError(f"スクリーンショット取得に失敗しました: モニター[{monitor_index}]: {str(e)}")

    def _run_all_steps(self):
        """ステップをループ実行"""
        try:
            self._execute_loop(self._run_steps_for_monitor, "ループ")
            logger.info("ステップのループ処理を開始しました")
        except Exception as e:
            logger.error(f"ステップループ実行エラー: {e}")
            self.update_status("❌ エラー発生")
            self.show_error_with_sound("エラー", f"ステップのループ実行に失敗しました: {e}")

    def _run_steps_all_monitors(self):
        """全モニターでステップを検索実行"""
        try:
            self._execute_loop(self._run_all_monitors, "全モニター検索")
            logger.info("全モニター検索のループ処理を開始しました")
        except Exception as e:
            logger.error(f"全モニター検索ループエラー: {e}")
            self.update_status("❌ エラー発生")
            self.show_error_with_sound("エラー", f"全モニター検索のループ実行に失敗しました: {e}")

    def _execute_loop(self, run_func: callable, mode: str):
        """ループ実行の共通ロジック"""
        try:
            loop_count = self.loop_count if self.loop_count > 0 else float('inf')
            for i in range(int(loop_count)):
                if not self.running:
                    logger.info(f"{mode} 実行が中断されました")
                    break
                self.update_status(f"🔄 {mode} {i + 1}/{self.loop_count if self.loop_count else '∞'}")
                logger.info(f"{mode} 実行中: ループ {i + 1}/{self.loop_count or '∞'}")
                run_func()
            # 実行完了時にボタン状態を更新
            self.running = False
            self.update_execution_buttons(False)
            
            status_text = "✅ 実行完了"
            if hasattr(self, 'main_status_label'):
                self.main_status_label.config(text=status_text)
            elif hasattr(self, 'status_label'):
                self.status_label.config(text=status_text)
            self.execution_start_index = 0  # リセット
            logger.info(f"{mode} 実行完了")
            
            # 処理完了通知ダイアログ（通知音付き）
            self.show_completion_notification()
        except Exception as e:
            logger.error(f"{mode} ループ実行エラー: {e}")
            self.update_status("❌ エラー発生")
            self.stop_execution()
            self.show_error_with_sound("エラー", f"{mode}の実行中にエラーが発生しました: {e}")

    def _run_steps_for_monitor(self):
        """指定モニターでステップを実行"""
        try:
            result = self._execute_steps_for_monitor(self.selected_monitor)
            logger.info(f"モニター[{self.selected_monitor}]のステップ実行: 結果={result}")
            return result
        except Exception as e:
            logger.error(f"モニター[{self.selected_monitor}]ステップ実行エラー: {e}")
            raise RuntimeError(f"モニター[{self.selected_monitor}]のステップ実行に失敗: {e}")

    def _run_all_monitors(self):
        """全モニターを順に実行"""
        try:
            for monitor_index in range(len(self.monitors)):
                if not self.running:
                    logger.info("全モニター実行が中断されました")
                    break
                self.update_status(f"🔍 モニター[{monitor_index}] 検索中...")
                logger.info(f"モニター[{monitor_index}] 検索開始")
                self.selected_monitor = monitor_index
                self.select_monitor(str(monitor_index))
                if not self._execute_steps_for_monitor(monitor_index):
                    logger.info(f"モニター[{monitor_index}] 実行中断")
                    break
            logger.info("全モニターの実行が完了しました")
        except Exception as e:
            logger.error(f"全モニター実行エラー: {e}")
            self.update_status("❌ エラー発生")
            self.show_error_with_sound("エラー", f"全モニターの実行中にエラーが発生しました: {e}")

    def _execute_steps_for_monitor(self, monitor_index: int) -> bool:
        """モニターごとのステップ実行（繰り返しアクション対応）"""
        try:
            # monitor_indexが文字列の場合は整数に変換
            if isinstance(monitor_index, str):
                monitor_index = int(monitor_index)
            # 繰り返しアクション解析によるステップ実行計画生成
            execution_plan = self._generate_execution_plan()
            
            # 通常実行: 実行計画全体を使用
            logger.info(f"実行計画生成完了: 実行計画長={len(execution_plan)}")
            
            # 有効ステップの総数を計算（repeat_start/repeat_endも含む）
            total_valid_steps = sum(1 for step_index, _ in execution_plan 
                                   if self.steps[step_index].enabled)
            
            # 実行済みステップのカウンター
            executed_steps = 0
            
            for exec_index, (step_index, repeat_iter) in enumerate(execution_plan, start=1):
                if not self.running:
                    self.update_status("⛔ 実行中断")
                    logger.info("ステップ実行が中断されました")
                    return False
                
                step = self.steps[step_index]
                
                # 無効なステップをスキップ
                if not step.enabled:
                    logger.info(f"ステップスキップ: 行番号={step_index+1}, type={step.type} (無効)")
                    continue
                
                # 繰り返し表示
                repeat_text = f" (繰り返し{repeat_iter+1}回目)" if repeat_iter > 0 else ""
                self.update_status(f"▶️ ステップ {step_index+1}/{len(self.steps)}: {self.get_type_display(step)}{repeat_text}")
                logger.info(f"ステップ実行開始: 行番号={step_index+1}, type={step.type}, repeat_iter={repeat_iter}")
                
                # 有効ステップの場合、進捗を更新
                if step.enabled:
                    executed_steps += 1
                    # 進行状況を更新
                    self.progress_var.set(f"{executed_steps}/{total_valid_steps}")
                
                # 実行中のステップをハイライト表示
                self.highlight_current_step(step_index)
                
                # プログレス可視化を更新
                current_progress = executed_steps / total_valid_steps if total_valid_steps > 0 else 0
                step_name = f"{step.comment}" if step.comment else f"{step.type.upper()}"
                self.update_realtime_info(step_name, current_progress)
                self.update_execution_stats(step_name)
                
                try:
                    if step.type == "repeat_start":
                        # 繰り返し開始は実行ログのみ
                        logger.info(f"繰り返し開始: {step.params['count']}回")
                    elif step.type == "repeat_end":
                        # 繰り返し終了は実行ログのみ
                        logger.info(f"繰り返し終了")
                    elif step.type == "image_click":
                        self._execute_image_click(step, monitor_index)
                    elif step.type == "coord_click":
                        self._execute_coord_click(step)
                    elif step.type == "coord_drag":
                        self._execute_coord_drag(step)
                    elif step.type == "image_relative_right_click":
                        self._execute_image_right_click(step, monitor_index)
                    elif step.type == "sleep":
                        self._execute_sleep(step)
                    elif step.type == "custom_text":
                        self._execute_custom_text(step)
                    elif step.type == "cmd_command":
                        self._execute_cmd_command(step)
                    else:
                        self._execute_key_action(step)
                    logger.info(f"ステップ実行成功: 行番号={step_index+1}, step={step}")
                    
                    # 成功時の統計とアニメーション更新
                    self.update_execution_stats(success=True)
                    self.animate_step_completion(step_index, success=True)
                except Exception as e:
                    self.update_status(f"❌ エラー発生: 行番号 {step_index+1}")
                    logger.error(f"ステップ実行エラー: 行番号={step_index+1}, step={step}, error={str(e)}")
                    
                    # スマートエラー分析とダイアログ表示
                    analysis = self.analyze_error(e, step, step_index+1)
                    self.show_error_dialog(analysis, step, step_index+1)
                    
                    
                    self.running = False
                    self.update_execution_buttons(False)
                    # Select the erroneous step in the Treeview
                    try:
                        children = self.tree.get_children()
                        if step_index < len(children):  # step_index is 0-based, children is 0-based
                            self.tree.selection_set(children[step_index])
                            self.tree.see(children[step_index])  # Ensure the selected item is visible
                            self.update_image_preview(step_index)  # 画像プレビューも同期
                            logger.info(f"エラー行を選択: 行番号={step_index+1}")
                        else:
                            logger.warning(f"エラー行の選択に失敗: 行番号={step_index+1}, ツリーアイテム数={len(children)}")
                    except Exception as select_error:
                        logger.error(f"エラー行の選択エラー: 行番号={step_index+1}, error={str(select_error)}")
                    # エラー統計を更新
                    self.update_execution_stats(error=True)
                    self.animate_step_completion(step_index, success=False)
                    
                    # エラー時にハイライトをクリア（コメントアウト - エラー行を視認しやすくするため）
                    # self.clear_current_step_highlight()
                    return False
            # 完了時の進捗表示を更新（100%完了）
            self.progress_var.set(f"{total_valid_steps}/{total_valid_steps}")
            self.update_realtime_info("全ステップ実行完了", 1.0)
            if hasattr(self, 'main_progress_bar') and self.main_progress_bar:
                self.update_progress_bar(self.main_progress_bar, 1.0, "全ステップ実行完了 ✅", animate=True)
                
            logger.info(f"モニター[{monitor_index}]の全ステップ実行完了")
            return True
        except Exception as e:
            self.update_status("❌ エラー発生")
            logger.error(f"モニター[{monitor_index}]ステップ実行エラー: {e}")
            self.show_error_with_sound("エラー", f"モニター[{monitor_index}]のステップ実行に失敗しました: 行番号なし, エラー: {e}")
            self.running = False
            self.update_execution_buttons(False)
            # エラー時にハイライトをクリア（コメントアウト - エラー行を視認しやすくするため）
            # self.clear_current_step_highlight()
            return False

    def _generate_execution_plan(self):
        """繰り返しアクション解析による実行計画生成（ネスト対応）"""
        return self._expand_nested_loops(self.steps, 0, len(self.steps))

    def _expand_nested_loops(self, steps, start_idx, end_idx, nest_level=0):
        """ネストしたループを再帰的に展開"""
        execution_plan = []
        i = start_idx
        
        while i < end_idx:
            step = steps[i]
            
            if step.type == "repeat_start":
                repeat_count = step.params.get('count', 1)
                execution_plan.append((i, nest_level))  # repeat_start自体を追加
                
                # 対応するrepeat_endを見つける
                nest_depth = 1
                end_pos = i + 1
                while end_pos < end_idx and nest_depth > 0:
                    if steps[end_pos].type == "repeat_start":
                        nest_depth += 1
                    elif steps[end_pos].type == "repeat_end":
                        nest_depth -= 1
                    end_pos += 1
                
                if nest_depth == 0:  # 対応するrepeat_endが見つかった
                    end_pos -= 1  # repeat_endのインデックスに調整
                    
                    # 繰り返し処理
                    for repeat_iter in range(repeat_count):
                        # ループ内容を再帰的に展開
                        inner_plan = self._expand_nested_loops(
                            steps, i + 1, end_pos, nest_level + repeat_iter + 1
                        )
                        execution_plan.extend(inner_plan)
                        # 各繰り返しの最後にrepeat_endも実行
                        execution_plan.append((end_pos, nest_level + repeat_iter + 1))
                    
                    i = end_pos + 1  # repeat_endの次に進む
                else:
                    raise ValueError(f"対応するrepeat_endが見つかりません: ステップ{i}")
                    
            elif step.type == "repeat_end":
                # 単体のrepeat_endは無視（親の処理で対応済み）
                i += 1
            else:
                # 通常のステップ
                execution_plan.append((i, nest_level))
                i += 1
        
        return execution_plan

    def _execute_image_click(self, step: Step, monitor_index: int):
        """画像をクリックまたはダブルクリック"""
        # monitor_indexが文字列の場合は整数に変換
        if isinstance(monitor_index, str):
            monitor_index = int(monitor_index)
            
        params = step.params
        try:
            path = params["path"]
            threshold = float(params.get("threshold"))
            click_type = params.get("click_type", "single")
            retry = int(params["retry"])
            delay = float(params["delay"])
            logger.info(f"画像クリック実行: path={path}, monitor={monitor_index}, threshold={threshold}, click_type={click_type}, retry={retry}, delay={delay}")

            template = cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)
            if template is None:
                raise ValueError(f"画像ファイルの読み込みに失敗しました: {path}")

            min_x, min_y, total_w, total_h = self.get_monitor_region(monitor_index)

            for attempt in range(retry + 1):
                if not self.running:
                    logger.info(f"画像クリック実行中断: attempt={attempt}")
                    break
                try:
                    screenshot = self.capture_screenshot(monitor_index)
                except RuntimeError as e:
                    logger.error(f"スクリーンショット取得エラー: {e}")
                    raise e

                result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(result)

                if max_val >= threshold:
                    h, w = template.shape[:2]
                    scale = self.dpi_scale / 100.0
                    click_point = (
                        int(min_x + (max_loc[0] + w // 2) * scale),
                        int(min_y + (max_loc[1] + h // 2) * scale),
                    )
                    if click_type == "single":
                        pyautogui.click(click_point)
                        logger.info(f"シングルクリック成功: point={click_point}")
                    elif click_type == "double":
                        pyautogui.doubleClick(click_point)
                        logger.info(f"ダブルクリック成功: point={click_point}")
                    elif click_type == "right":
                        pyautogui.click(click_point, button='right')
                        logger.info(f"右クリック成功: point={click_point}")
                    else:
                        raise ValueError(f"無効なクリックタイプ: {click_type}")
                    return
                time.sleep(delay)
                logger.info(f"画像検索試行: attempt={attempt + 1}, max_val={max_val}")

            # 画像が見つからない場合、例外を投げて実行を停止
            raise RuntimeError(f"画像が見つかりませんでした: {os.path.basename(path)}")
        except Exception as e:
            logger.error(f"画像クリック実行エラー: path={path}, error={str(e)}")
            raise RuntimeError(f"画像クリックの実行に失敗しました: {e}")



    def _execute_coord_click(self, step: Step):
        """指定座標でクリック（シングル・ダブル・右クリック対応）"""
        params = step.params
        try:
            x = int(params["x"])
            y = int(params["y"])
            click_type = params.get("click_type", "single")
            
            logger.info(f"座標{click_type}クリック実行: point=({x}, {y})")
            
            if click_type == "single":
                pyautogui.click(x, y)
            elif click_type == "double":
                pyautogui.doubleClick(x, y)
            elif click_type == "right":
                pyautogui.rightClick(x, y)
            else:
                # フォールバック：デフォルトはシングルクリック
                pyautogui.click(x, y)
                
            logger.info(f"座標{click_type}クリック成功: point=({x}, {y})")
        except Exception as e:
            click_type = params.get("click_type", "single")
            logger.error(f"座標{click_type}クリックエラー: point=({x}, {y}), error={str(e)}")
            raise RuntimeError(f"座標{click_type}クリックの実行に失敗しました: {e}")

    def _execute_coord_drag(self, step: Step):
        """座標間ドラッグを実行"""
        params = step.params
        try:
            start_x = int(params["start_x"])
            start_y = int(params["start_y"])
            end_x = int(params["end_x"])
            end_y = int(params["end_y"])
            duration = float(params["duration"])
            
            logger.info(f"座標ドラッグ実行: start=({start_x}, {start_y}), end=({end_x}, {end_y}), duration={duration}秒")
            
            # 開始座標に移動してからドラッグ実行
            pyautogui.moveTo(start_x, start_y)
            pyautogui.dragTo(end_x, end_y, duration, button='left')
                
            logger.info(f"座標ドラッグ成功: start=({start_x}, {start_y}), end=({end_x}, {end_y})")
        except Exception as e:
            logger.error(f"座標ドラッグエラー: start=({start_x}, {start_y}), end=({end_x}, {end_y}), error={str(e)}")
            raise RuntimeError(f"座標ドラッグの実行に失敗しました: {e}")


    def _execute_image_right_click(self, step: Step, monitor_index: int):
        """画像のオフセット座標でクリック"""
        # monitor_indexが文字列の場合は整数に変換
        if isinstance(monitor_index, str):
            monitor_index = int(monitor_index)
            
        params = step.params
        try:
            path = params["path"]
            threshold = float(params["threshold"])
            click_type = params.get("click_type", "right")
            offset_x = int(params["offset_x"])
            offset_y = int(params["offset_y"])
            retry = int(params["retry"])
            delay = float(params["delay"])
            logger.info(f"画像オフセット{click_type}クリック実行: path={path}, monitor={monitor_index}, threshold={threshold}, offset=({offset_x}, {offset_y}), retry={retry}, delay={delay}")

            template = cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)
            if template is None:
                raise ValueError(f"画像ファイルの読み込みに失敗しました: {path}")

            min_x, min_y, total_w, total_h = self.get_monitor_region(monitor_index)

            for attempt in range(retry + 1):
                if not self.running:
                    logger.info(f"画像右クリック実行中断: attempt={attempt}")
                    break
                try:
                    screenshot = self.capture_screenshot(monitor_index)
                except RuntimeError as e:
                    logger.error(f"スクリーンショット取得エラー: {e}")
                    raise e

                result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(result)

                if max_val >= threshold:
                    h, w = template.shape[:2]
                    scale = self.dpi_scale / 100.0
                    base_point = (
                        int(min_x + (max_loc[0] + w // 2) * scale),
                        int(min_y + (max_loc[1] + h // 2) * scale),
                    )
                    click_point = (
                        base_point[0] + int(offset_x * scale),
                        base_point[1] + int(offset_y * scale),
                    )
                    
                    if click_type == "single":
                        pyautogui.click(click_point)
                    elif click_type == "double":
                        pyautogui.doubleClick(click_point)
                    elif click_type == "right":
                        pyautogui.rightClick(click_point)
                    else:
                        raise ValueError(f"無効なクリックタイプ: {click_type}")
                    
                    logger.info(f"画像オフセット{click_type}クリック成功: point={click_point}")
                    return
                time.sleep(delay)
                logger.info(f"画像検索試行: attempt={attempt + 1}, max_val={max_val}")

            # 画像が見つからない場合、例外を投げて実行を停止
            raise RuntimeError(f"画像が見つかりませんでした: {os.path.basename(path)}")
        except Exception as e:
            logger.error(f"画像右クリック実行エラー: path={path}, error={str(e)}")
            raise RuntimeError(f"画像右クリックの実行に失敗しました: {e}")



    def _execute_key_action(self, step: Step):
        """キー操作を実行"""
        params = step.params
        try:
            key = params["key"]
            logger.info(f"キー操作実行: key={key}")
            
            # スクリーンショット専用処理
            if key.lower() == "ctrl+shift+f12":
                self.take_screenshot_and_save()
                logger.info("スクリーンショット機能実行完了")
                return
            
            if "+" in key:
                keys = [k.strip().lower() for k in key.split("+")]
                pyautogui.hotkey(*keys)
                logger.info(f"ホットキー成功: keys={keys}")
            else:
                pyautogui.press(key.lower())
                logger.info(f"キー押下成功: key={key}")
        except Exception as e:
            logger.error(f"キー操作エラー: key={key}, error={str(e)}")
            raise RuntimeError(f"キー操作の実行に失敗しました: {e}")

    def _execute_custom_text(self, step: Step):
            """カスタム文字列を入力"""
            params = step.params
            try:
                text = params["text"]
                logger.info(f"カスタム文字列入力実行: text={text}")
                
                # Copy text to clipboard and paste it
                pyperclip.copy(text)
                time.sleep(0.1)  # Brief pause to ensure clipboard is updated
                pyautogui.hotkey("ctrl", "v")
                time.sleep(0.1)  # Brief pause to ensure paste completes

                logger.info(f"カスタム文字列入力成功: text={text}")
            except Exception as e:
                logger.error(f"カスタム文字列入力エラー: text={text}, error={str(e)}")
                raise RuntimeError(f"カスタム文字列入力の実行に失敗しました: {e}")

    def _validate_command_safety(self, command: str) -> bool:
        """コマンドの安全性を検証"""
        # Windows cmd/PowerShell の危険なコマンドパターン
        dangerous_patterns = [
            # ファイル削除系
            r'del\s+.*(/[sq]|\\|\*)',           # del /s, del /q, del *.* 
            r'rmdir\s+.*(/s|/q)',               # rmdir /s
            r'rd\s+.*(/s|/q)',                  # rd /s /q
            r'erase\s+.*(\*|\\)',               # erase
            r'remove-item\s+.*(-recurse|-force)', # PowerShell Remove-Item
            
            # システム操作系
            r'format\s+[a-z]:',                 # format C:
            r'fdisk\s+',                        # fdisk
            r'diskpart\s*$',                    # diskpart
            r'bootrec\s+',                      # bootrec
            r'bcdedit\s+',                      # bcdedit
            
            # システム制御系
            r'shutdown\s+.*(/[rs]|/[fth])',     # shutdown /r /s /f /t /h
            r'restart-computer',                # PowerShell restart
            r'stop-computer',                   # PowerShell shutdown
            
            # ネットワーク/セキュリティ系
            r'netsh\s+',                        # netsh (firewall, wifi等)
            r'netstat\s+.*(-a|-n)',             # netstat
            r'arp\s+(-[ads])',                  # arp manipulation
            
            # レジストリ操作系
            r'reg\s+(delete|add|import)',       # registry operations
            r'regedit\s+(/[si])',              # regedit import/silent
            
            # サービス制御系  
            r'sc\s+(delete|create|config)',     # service control
            r'net\s+(start|stop|user)',         # net commands
            r'wmic\s+',                         # wmic
            
            # プロセス制御系
            r'taskkill\s+.*(/f|/im)',          # force kill processes
            r'tskill\s+',                       # tskill
            r'stop-process\s+.*-force',         # PowerShell force stop
            
            # PowerShell危険系
            r'powershell\s.*(-encodedcommand|-enc|-ep\s+bypass)', # encoded/bypass execution policy
            r'invoke-expression\s*\(',          # Invoke-Expression
            r'iex\s*\(',                        # iex alias
            r'invoke-webrequest.*downloadfile', # file download
            r'start-process\s+.*-windowstyle\s+hidden', # hidden process
            
            # スケジュール/自動実行系
            r'schtasks\s+.*(/create|/delete)',  # scheduled tasks
            r'at\s+\d+:\d+',                   # at command
            
            # ファイアウォール/セキュリティ系
            r'netsh\s+advfirewall',            # firewall
            r'netsh\s+firewall',               # legacy firewall
            
            # コマンドチェーン（危険な組み合わせ）
            r'[;&|`]\s*(del|rmdir|rd|format|shutdown)', # command chaining
            r'\|\s*(del|rmdir|rd|format)',     # pipe to dangerous commands
            
            # 隠蔽系
            r'attrib\s+.*\+[hs]',              # hide files
            r'icacls\s+.*(/deny|/remove)',     # permission manipulation
            
            # バッチ/スクリプト実行
            r'cmd\s+/c\s+.*(\||&)',            # cmd /c with chaining
            r'start\s+/min',                   # minimized start
        ]
        
        import re
        command_lower = command.lower()
        
        for pattern in dangerous_patterns:
            if re.search(pattern, command_lower):
                return False
        
        # 基本的な文字検証（制御文字やnull文字を除外）
        if any(ord(c) < 32 and c not in ['\t', '\n', '\r'] for c in command):
            return False
            
        return True

    def _execute_cmd_command(self, step: Step):
        """cmdコマンドを実行"""
        params = step.params
        try:
            command = params["command"]
            timeout = params.get("timeout", 30)
            wait_completion = params.get("wait_completion", True)
            scheduled_time = params.get("scheduled_time", None)
            
            # コマンド安全性チェック
            if not self._validate_command_safety(command):
                error_msg = f"安全でないコマンドが検出されました: {command[:100]}..."
                logger.error(f"危険なコマンド実行を拒否: {command}")
                raise RuntimeError(error_msg)
            
            logger.info(f"cmdコマンド実行: command={command}, timeout={timeout}, wait_completion={wait_completion}, scheduled_time={scheduled_time}")
            
            # 時間指定がある場合は待機（日次繰り返し対応）
            if scheduled_time:
                if self._wait_for_scheduled_time(scheduled_time):
                    return  # 実行が中断された場合
            
            self.update_status(f"🔧 コマンド実行中: {command[:50]}...")
            
            if wait_completion:
                # 完了を待つ場合
                result = subprocess.run(
                    command,
                    shell=True,
                    capture_output=True,
                    text=True,
                    timeout=timeout,
                    encoding='utf-8',
                    errors='replace'
                )
                
                if result.returncode == 0:
                    logger.info(f"cmdコマンド実行成功: command={command}, stdout={result.stdout}")
                    if result.stdout.strip():
                        messagebox.showinfo("コマンド実行結果", f"コマンド: {command[:50]}...\n\n実行結果:\n{result.stdout}")
                else:
                    # よくあるエラーの場合はよりわかりやすいメッセージにする
                    if "No such file or directory" in result.stderr or "cannot find" in result.stderr.lower():
                        error_msg = f"ファイルまたはディレクトリが見つかりません\n\nコマンド: {command}\n\nエラー詳細:\n{result.stderr}"
                    elif "command not found" in result.stderr.lower() or "is not recognized" in result.stderr.lower():
                        error_msg = f"コマンドが見つかりません\n\nコマンド: {command}\n\nエラー詳細:\n{result.stderr}"
                    else:
                        error_msg = f"コマンドが失敗しました (終了コード: {result.returncode})\n\nコマンド: {command}\n\nエラー詳細:\n{result.stderr}"
                    
                    logger.error(f"cmdコマンド実行失敗: command={command}, stderr={result.stderr}")
                    raise RuntimeError(error_msg)
            else:
                # バックグラウンドで実行（完了を待たない）
                # 同じ安全性チェックが適用済み
                subprocess.Popen(
                    command,
                    shell=True,
                    creationflags=subprocess.CREATE_NEW_CONSOLE if os.name == 'nt' else 0
                )
                logger.info(f"cmdコマンドをバックグラウンドで開始: command={command}")
                
        except subprocess.TimeoutExpired:
            error_msg = f"コマンドがタイムアウトしました ({timeout}秒)"
            logger.error(f"cmdコマンドタイムアウト: command={command}, timeout={timeout}")
            raise RuntimeError(error_msg)
        except Exception as e:
            logger.error(f"cmdコマンド実行エラー: command={command}, error={str(e)}")
            raise RuntimeError(f"cmdコマンドの実行に失敗しました: {e}")

    def _wait_for_scheduled_time(self, scheduled_time: str) -> bool:
        """指定時刻まで待機（日次繰り返し対応）
        
        Args:
            scheduled_time: HH:MM:SS 形式の時刻文字列
            
        Returns:
            bool: True if execution was interrupted, False if continued
        """
        try:
            # HH:MM:SS形式のパース
            time_parts = scheduled_time.split(':')
            if len(time_parts) != 3:
                raise ValueError("時刻の形式が正しくありません。HH:MM:SS形式で入力してください。")
            
            target_hour = int(time_parts[0])
            target_minute = int(time_parts[1])
            target_second = int(time_parts[2])
            
            # 現在時刻取得
            now = datetime.now()
            
            # 今日の指定時刻を作成
            target_today = now.replace(hour=target_hour, minute=target_minute, second=target_second, microsecond=0)
            
            # 既に今日の指定時刻を過ぎている場合は明日の同時刻に設定
            if target_today <= now:
                target_today += timedelta(days=1)
            
            wait_seconds = (target_today - now).total_seconds()
            logger.info(f"指定時刻まで待機: {wait_seconds}秒 (目標時刻: {target_today.strftime('%Y-%m-%d %H:%M:%S')})")
            
            # 待機中のステータス表示
            self.update_status(f"⏰ {target_today.strftime('%H:%M:%S')}まで待機中...")
            
            # 1秒ごとにチェックしながら待機
            remaining = int(wait_seconds)
            while remaining > 0 and self.running:
                if remaining % 60 == 0:  # 1分ごとにログ出力
                    logger.info(f"実行待機中: 残り{remaining // 60}分{remaining % 60}秒")
                
                # 残り時間を表示
                hours = remaining // 3600
                minutes = (remaining % 3600) // 60
                seconds = remaining % 60
                
                if hours > 0:
                    time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                else:
                    time_str = f"{minutes:02d}:{seconds:02d}"
                
                self.update_status(f"⏰ 実行待機中... あと{time_str}")
                
                time.sleep(1)
                remaining -= 1
            
            # 実行が中断された場合
            if not self.running:
                logger.info("時間待機が中断されました")
                return True
            
            logger.info(f"指定時刻に到達しました: {target_today.strftime('%H:%M:%S')}")
            return False
            
        except ValueError as e:
            logger.error(f"時刻フォーマットエラー: {scheduled_time}, error={str(e)}")
            raise RuntimeError(f"時刻の形式が正しくありません: {e}")
        except Exception as e:
            logger.error(f"時間待機エラー: {scheduled_time}, error={str(e)}")
            raise RuntimeError(f"時間待機に失敗しました: {e}")

    def _execute_sleep(self, step: Step):
        """スリープを実行"""
        params = step.params
        try:
            wait_type = params.get("wait_type", "sleep")  # デフォルトはスリープ
            
            if wait_type == "scheduled":
                # 時刻指定待機
                scheduled_time = params.get("scheduled_time")
                if scheduled_time:
                    logger.info(f"時刻指定待機実行: scheduled_time={scheduled_time}")
                    if self._wait_for_scheduled_time(scheduled_time):
                        return  # 実行が中断された場合
                    logger.info(f"時刻指定待機完了: scheduled_time={scheduled_time}")
                else:
                    logger.error("scheduled_timeパラメータが見つかりません")
                    raise RuntimeError("時刻指定待機のパラメータが見つかりません")
            else:
                # デフォルトのスリープ（秒数指定）
                seconds = float(params.get("seconds", 1.0))
                logger.info(f"スリープ実行: seconds={seconds}")
                time.sleep(seconds)
                logger.info(f"スリープ完了: seconds={seconds}")
        except Exception as e:
            logger.error(f"待機エラー: params={params}, error={str(e)}")
            raise RuntimeError(f"待機の実行に失敗しました: {e}")

    def analyze_error(self, error: Exception, step: Step, step_number: int) -> Dict[str, str]:
        """エラーを分析して詳細情報と対処法を提供"""
        error_str = str(error).lower()
        error_type = type(error).__name__
        
        analysis = {
            "type": error_type,
            "message": str(error),
            "step_info": f"行 {step_number}: {step.type}",
            "suggestion": "詳細ログを確認してください",
            "action": "retry"
        }
        
        # 画像関連エラー
        if "画像が見つかりません" in str(error):
            analysis.update({
                "category": "🔍 画像検索失敗",
                "reason": "指定した画像が画面上に見つかりませんでした",
                "suggestion": "• 画像の信頼度を下げる (推奨: 0.7-0.8)\n• 画像を再キャプチャする\n• 画面の状態を確認する",
                "action": "adjust_threshold"
            })
        elif "ファイルの読み込みに失敗" in str(error):
            analysis.update({
                "category": "📁 ファイルエラー",
                "reason": "画像ファイルが存在しないか、破損しています",
                "suggestion": "• ファイルパスを確認する\n• 画像を再選択する\n• ファイル権限を確認する",
                "action": "reselect_file"
            })
        # キー操作エラー  
        elif "キー操作" in str(error):
            analysis.update({
                "category": "⌨ キーボードエラー",
                "reason": "キー入力の実行に失敗しました",
                "suggestion": "• キーの組み合わせを確認する\n• アプリケーションがアクティブか確認する\n• 短い待機時間を追加する",
                "action": "check_focus"
            })
        # 座標エラー
        elif "座標" in str(error) or "click" in error_str:
            analysis.update({
                "category": "🖱 マウスエラー", 
                "reason": "マウス操作の実行に失敗しました",
                "suggestion": "• 座標値を確認する\n• 画面解像度の変更がないか確認する\n• ウィンドウ位置を確認する",
                "action": "recapture_coords"
            })
        # スクリーンショットエラー
        elif "スクリーンショット" in str(error):
            analysis.update({
                "category": "📷 画面キャプチャエラー",
                "reason": "画面のキャプチャに失敗しました",
                "suggestion": "• モニター設定を確認する\n• アプリケーションを管理者権限で実行する\n• DPI設定を確認する",
                "action": "check_permissions"
            })
        
        return analysis

    def show_error_with_sound(self, title: str, message: str):
        """エラー音付きのエラーダイアログ表示"""
        # エラー音を再生
        try:
            import winsound
            winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
        except ImportError:
            # winsoundが使えない場合はbeep音
            try:
                import os
                os.system("echo \a")  # システムbeep音
            except:
                pass  # 音が出せない場合は無視
        
        # メッセージボックスを最前面で表示
        self.root.attributes("-topmost", True)

    def show_completion_notification(self):
        """処理完了通知ダイアログ（通知音付き）"""
        try:
            # 完了音を再生
            try:
                import winsound
                # 成功音を1回再生
                winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)
            except ImportError:
                # winsoundが使えない場合はbeep音
                try:
                    self.root.bell()
                except:
                    pass
            
            # メッセージボックスを最前面で表示
            self.root.attributes("-topmost", True)
            messagebox.showinfo("実行完了", "✅ 処理が完了しました。")
            self.root.attributes("-topmost", False)
            logger.info("処理完了通知を表示しました")
        except Exception as e:
            logger.error(f"完了通知表示失敗: {e}")
            # フォールバック：シンプルな表示
            try:
                messagebox.showinfo("実行完了", "処理が完了しました。")
            except:
                logger.error("完了通知ダイアログの表示に失敗しました")
    
    def show_error_dialog(self, analysis: Dict[str, str], step: Step, step_number: int):
        """詳細なエラーダイアログを表示"""
        # エラー音を再生
        try:
            import winsound
            winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
        except ImportError:
            # winsoundが使えない場合はbeep音
            try:
                import os
                os.system("echo \a")  # システムbeep音
            except:
                pass  # 音が出せない場合は無視
        
        dialog = tk.Toplevel(self.root)
        dialog.title("🚨 実行エラー詳細")
        dialog.geometry("500x450")  # Width reduced to maintain proportion
        dialog.configure(bg="#2b2b2b")
        dialog.resizable(False, False)
        dialog.grab_set()
        
        # ウィンドウを最前面に表示
        dialog.attributes("-topmost", True)
        dialog.focus_force()
        dialog.lift()
        
        # ウィンドウを中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        main_frame = tk.Frame(dialog, bg="#2b2b2b")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # エラーカテゴリ
        title_label = tk.Label(
            main_frame,
            text=analysis.get("category", "❌ エラー"),
            font=("Meiryo UI", 16, "bold"),
            fg="#ff6b6b",
            bg="#2b2b2b"
        )
        title_label.pack(anchor="w", pady=(0, 10))
        
        # ステップ情報
        step_frame = tk.Frame(main_frame, bg="#3c3c3c", relief="groove", bd=1)
        step_frame.pack(fill="x", pady=(0, 10))
        
        step_info_label = tk.Label(
            step_frame,
            text=f"📍 {analysis['step_info']}",
            font=("Meiryo UI", 10, "bold"),
            fg="#ffd93d", 
            bg="#3c3c3c"
        )
        step_info_label.pack(anchor="w", padx=10, pady=5)
        
        # エラー詳細
        details_frame = tk.Frame(main_frame, bg="#2b2b2b")
        details_frame.pack(fill="both", expand=True)
        
        # 原因
        reason_label = tk.Label(
            details_frame,
            text="🔍 エラーの原因:",
            font=("Meiryo UI", 12, "bold"),
            fg="#74c0fc",
            bg="#2b2b2b"
        )
        reason_label.pack(anchor="w", pady=(0, 5))
        
        reason_text = tk.Text(
            details_frame,
            height=3,
            bg="#3c3c3c",
            fg="#ffffff",
            font=("Meiryo UI", 10),
            wrap="word",
            relief="flat",
            padx=10,
            pady=5
        )
        reason_text.pack(fill="x", pady=(0, 10))
        reason_text.insert("1.0", analysis.get("reason", "不明なエラーが発生しました"))
        reason_text.config(state="disabled")
        
        # 対処法
        solution_label = tk.Label(
            details_frame,
            text="💡 推奨される対処法:",
            font=("Meiryo UI", 12, "bold"),
            fg="#51cf66",
            bg="#2b2b2b"
        )
        solution_label.pack(anchor="w", pady=(0, 5))
        
        solution_text = tk.Text(
            details_frame,
            height=6,
            bg="#3c3c3c",
            fg="#ffffff", 
            font=("Meiryo UI", 10),
            wrap="word",
            relief="flat",
            padx=10,
            pady=5
        )
        solution_text.pack(fill="both", expand=True, pady=(0, 10))
        solution_text.insert("1.0", analysis.get("suggestion", "ログファイルを確認してください"))
        solution_text.config(state="disabled")
        
        # ボタンフレーム
        button_frame = tk.Frame(main_frame, bg="#2b2b2b")
        button_frame.pack(fill="x", pady=(10, 0))
        
        
        edit_btn = tk.Button(
            button_frame,
            text="✏ 設定を編集",
            command=lambda: self.edit_step_from_error(step, step_number, dialog),
            font=("Meiryo UI", 10, "bold"),
            bg="#495057",
            fg="white",
            relief="flat",
            padx=15,
            pady=5,
            cursor="hand2"
        )
        edit_btn.pack(side="left", padx=(0, 10))
        
        close_btn = tk.Button(
            button_frame,
            text="❌ 閉じる",
            command=dialog.destroy,
            font=("Meiryo UI", 10, "bold"),
            bg="#6c757d",
            fg="white",
            relief="flat",
            padx=15,
            pady=5,
            cursor="hand2"
        )
        close_btn.pack(side="right")

    def edit_step_from_error(self, step: Step, step_number: int, dialog: tk.Toplevel):
        """エラーダイアログからステップ編集"""
        dialog.destroy()
        # 該当ステップを選択
        children = self.tree.get_children()
        if step_number - 1 < len(children):
            item_id = children[step_number - 1]
            self.tree.selection_set(item_id)
            self.tree.see(item_id)
            # 編集モードを開く
            self.edit_selected_step()

    
    def save_state(self, action_description: str):
        """現在の状態をUndo stackに保存"""
        try:
            state = {
                'id': str(uuid.uuid4()),
                'timestamp': datetime.now().isoformat(),
                'action': action_description,
                'steps': [asdict(step) for step in self.steps],
                'monitor_index': self.monitor_var.get()
            }
            
            # 新しい操作が行われた場合、redo stackをクリア
            self.redo_stack.clear()
            
            # 現在の状態をundo stackに保存
            if self.current_state_id:
                self.undo_stack.append(state)
                
            self.current_state_id = state['id']
            logger.info(f"状態保存: {action_description} (ID: {state['id'][:8]})")
            
        except Exception as e:
            logger.error(f"状態保存エラー: {e}")
    
    def undo(self):
        """Undo操作を実行"""
        try:
            if not self.undo_stack:
                messagebox.showinfo("Undo", "元に戻すことのできる操作がありません")
                return
                
            # 現在の状態をredo stackに保存
            current_state = {
                'id': str(uuid.uuid4()),
                'timestamp': datetime.now().isoformat(),
                'action': 'Current State',
                'steps': [asdict(step) for step in self.steps],
                'monitor_index': self.monitor_var.get()
            }
            self.redo_stack.append(current_state)
            
            # 前の状態を復元
            previous_state = self.undo_stack.pop()
            self.restore_state(previous_state)
            
            # ステータス表示
            action = previous_state['action']
            self.update_status(f"↩️ 元に戻しました: {action}")
            logger.info(f"Undo実行: {action}")
            
        except Exception as e:
            logger.error(f"Undoエラー: {e}")
            self.show_error_with_sound("エラー", f"Undo操作に失敗しました: {e}")
    
    def redo(self):
        """Redo操作を実行"""
        try:
            if not self.redo_stack:
                messagebox.showinfo("Redo", "やり直すことのできる操作がありません")
                return
                
            # 現在の状態をundo stackに保存
            current_state = {
                'id': str(uuid.uuid4()),
                'timestamp': datetime.now().isoformat(),
                'action': 'Current State',
                'steps': [asdict(step) for step in self.steps],
                'monitor_index': self.monitor_var.get()
            }
            self.undo_stack.append(current_state)
            
            # 次の状態を復元
            next_state = self.redo_stack.pop()
            self.restore_state(next_state)
            
            # ステータス表示
            action = next_state['action']
            self.update_status(f"↪️ やり直しました: {action}")
            logger.info(f"Redo実行: {action}")
            
        except Exception as e:
            logger.error(f"Redoエラー: {e}")
            self.show_error_with_sound("エラー", f"Redo操作に失敗しました: {e}")
    
    def restore_state(self, state: dict):
        """指定した状態を復元"""
        try:
            # ステップリストを復元
            self.steps.clear()
            for step_dict in state['steps']:
                step = Step(
                    step_dict['type'],
                    step_dict['params'],
                    step_dict['comment'],
                    step_dict.get('enabled', True)
                )
                self.steps.append(step)
            
            # モニター選択を復元
            self.monitor_var.set(state['monitor_index'])
            
            # UIを更新
            self.update_tree()
            self.current_state_id = state['id']
            
        except Exception as e:
            logger.error(f"状態復元エラー: {e}")
            raise


class MouseCoordinateDialog:
    """マウス座標選択ダイアログ（リアルタイム表示対応）"""
    
    def __init__(self, parent: tk.Tk, title: str = "座標選択"):
        self.parent = parent
        self.selected_coordinates = None
        self.tracking = False
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("700x600")  # クリック操作系で統一サイズに拡大
        self.dialog.configure(bg="#2b2b2b")
        self.dialog.resizable(True, True)  # リサイズ可能にしてボタンを確認可能に
        self.dialog.minsize(500, 400)
        # トラッキング時にiconify可能にするためtransientを削除
        # self.dialog.transient(parent)  # コメントアウト
        self.dialog.grab_set()
        
        # ウィンドウを中央に配置
        self.center_window()
        
        # UIを構築
        self.setup_ui()
        
        # ESCキーで閉じる
        self.dialog.bind('<Escape>', lambda e: self.close_dialog())
        
        # ウィンドウが閉じられた時のクリーンアップ
        self.dialog.protocol("WM_DELETE_WINDOW", self.close_dialog)
    
    def center_window(self):
        """ウィンドウをスクリーン中央に配置"""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f'+{x}+{y}')
    
    def setup_ui(self):
        """UIを構築"""
        main_frame = tk.Frame(self.dialog, bg="#2b2b2b")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # タイトル
        title_label = tk.Label(
            main_frame,
            text="🖱️ 座標を選択してください",
            font=("Meiryo UI", 14, "bold"),
            fg="#74c0fc",
            bg="#2b2b2b"
        )
        title_label.pack(pady=(0, 20))
        
        # 現在のマウス位置表示エリア
        position_frame = tk.Frame(main_frame, bg="#3c3c3c", relief="groove", bd=1)
        position_frame.pack(fill="x", pady=(0, 15))
        
        position_label = tk.Label(
            position_frame,
            text="📍 現在のマウス位置:",
            font=("Meiryo UI", 11, "bold"),
            fg="#51cf66",
            bg="#3c3c3c"
        )
        position_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        self.position_display = tk.Label(
            position_frame,
            text="X: ---- , Y: ----",
            font=("Meiryo UI", 16, "bold"),
            fg="#ffd93d",
            bg="#3c3c3c"
        )
        self.position_display.pack(pady=(0, 10))
        
        # トラッキング制御
        control_frame = tk.Frame(main_frame, bg="#2b2b2b")
        control_frame.pack(fill="x", pady=(0, 15))
        
        self.track_btn = tk.Button(
            control_frame,
            text="🎯 座標トラッキング開始",
            command=self.toggle_tracking,
            font=("Meiryo UI", 10, "bold"),
            bg="#0d7377",
            fg="white",
            relief="flat",
            padx=20,
            pady=8,
            cursor="hand2"
        )
        self.track_btn.pack()
        
        # 使用方法の説明
        instruction_label = tk.Label(
            main_frame,
            text="💡 使用方法:\n" +
                 "1. 「座標トラッキング開始」をクリック\n" +
                 "2. 目標の場所にマウスを移動\n" +
                 "3. 右クリックして座標を自動確定\n" +
                 "4. または手動で座標を入力して「座標を使用」",
            font=("Meiryo UI", 10),
            fg="#ffffff",
            bg="#2b2b2b",
            justify="left"
        )
        instruction_label.pack(pady=(0, 15))
        
        # 座標入力エリア
        input_frame = tk.Frame(main_frame, bg="#3c3c3c", relief="groove", bd=1)
        input_frame.pack(fill="x", pady=(0, 15))
        
        input_label = tk.Label(
            input_frame,
            text="✏️ 直接入力:",
            font=("Meiryo UI", 11, "bold"),
            fg="#ff6b6b",
            bg="#3c3c3c"
        )
        input_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        coords_frame = tk.Frame(input_frame, bg="#3c3c3c")
        coords_frame.pack(padx=10, pady=(0, 10))
        
        tk.Label(coords_frame, text="X:", font=("Meiryo UI", 10), fg="white", bg="#3c3c3c").pack(side="left")
        self.x_entry = tk.Entry(coords_frame, font=("Meiryo UI", 10), width=8)
        self.x_entry.pack(side="left", padx=(5, 15))
        
        tk.Label(coords_frame, text="Y:", font=("Meiryo UI", 10), fg="white", bg="#3c3c3c").pack(side="left")
        self.y_entry = tk.Entry(coords_frame, font=("Meiryo UI", 10), width=8)
        self.y_entry.pack(side="left", padx=(5, 0))
        
        # ボタンフレーム
        button_frame = tk.Frame(main_frame, bg="#2b2b2b")
        button_frame.pack(fill="x")
        
        # ttk.Buttonに変更して統一感を向上
        ok_btn = ttk.Button(
            button_frame,
            text="OK",
            command=self.confirm_coordinates,
            style='Primary.TButton',
            width=10
        )
        ok_btn.pack(side="right", padx=(10, 0))
        
        # ttk.Buttonに変更して統一感を向上
        cancel_btn = ttk.Button(
            button_frame,
            text="キャンセル",
            command=self.close_dialog,
            style='Modern.TButton',
            width=10
        )
        cancel_btn.pack(side="right")
    
    def toggle_tracking(self):
        """座標トラッキングの開始/停止"""
        if not self.tracking:
            self.start_tracking()
        else:
            self.stop_tracking()
    
    def start_tracking(self):
        """座標トラッキングを開始"""
        self.tracking = True
        self.track_btn.configure(
            text="🛑 トラッキング停止",
            bg="#ff6b6b"
        )
        
        # ダイアログを最小化せずに、透明度を下げて背景に
        try:
            self.dialog.attributes('-alpha', 0.7)  # 透明度を70%に調整（見やすく）
            self.dialog.attributes('-topmost', True)  # 最前面に固定
        except:
            # 透明度が使えない場合はそのまま
            pass
        
        # リアルタイム座標更新を開始
        self.update_position()
        
        # グローバル右クリック監視を開始
        self.start_global_click_detection()
    
    def stop_tracking(self):
        """座標トラッキングを停止"""
        self.tracking = False
        self.track_btn.configure(
            text="🎯 座標トラッキング開始",
            bg="#0d7377"
        )
        
        # ダイアログの表示を元に戻す
        try:
            self.dialog.attributes('-alpha', 1.0)  # 不透明に戻す
            self.dialog.attributes('-topmost', False)  # 最前面固定を解除
        except:
            pass
        
        # グローバルクリック監視を停止
        self.stop_global_click_detection()
    
    def update_position(self):
        """マウス位置を更新"""
        if self.tracking:
            try:
                # マウスの現在位置を取得
                x, y = pyautogui.position()
                self.position_display.configure(text=f"X: {x:4d} , Y: {y:4d}")
                
                # エントリーフィールドも更新
                self.x_entry.delete(0, "end")
                self.x_entry.insert(0, str(x))
                self.y_entry.delete(0, "end") 
                self.y_entry.insert(0, str(y))
                
                # 50ms後に再実行
                self.dialog.after(50, self.update_position)
            except Exception:
                self.stop_tracking()
    
    def on_click_coordinate(self, event=None):
        """座標確定クリック"""
        if self.tracking:
            try:
                x, y = pyautogui.position()
                self.selected_coordinates = (x, y)
                self.stop_tracking()
                
                # 確定音（Windowsシステムサウンド）
                import winsound
                winsound.MessageBeep(winsound.MB_OK)
                
                messagebox.showinfo("座標確定", f"座標が確定されました:\nX: {x}, Y: {y}")
            except Exception as e:
                logger.error(f"座標確定エラー: {e}")
    
    def confirm_coordinates(self):
        """座標を確定"""
        try:
            # エントリーフィールドから座標を取得
            x_text = self.x_entry.get()
            y_text = self.y_entry.get()
            
            x = int(x_text or "0")
            y = int(y_text or "0")
            
            self.selected_coordinates = (x, y)
            
            self.close_dialog()  # destroyではなくclose_dialogを使用
            
        except ValueError as e:
            self.show_error_dialog("エラー", "有効な座標を入力してください", parent=self.dialog)
        except Exception as e:
            self.show_error_dialog("エラー", f"座標確定に失敗しました: {e}", parent=self.dialog)
    
    def start_global_click_detection(self):
        """グローバル右クリック検出を開始"""
        try:
            import pynput.mouse as mouse
            
            def on_click(x, y, button, pressed):
                if pressed and button == mouse.Button.right and self.tracking:
                    # 右クリック検出
                    self.selected_coordinates = (x, y)
                    
                    # UI更新はメインスレッドで実行
                    self.dialog.after(0, lambda: self.on_right_click_detected(x, y))
                    return False  # リスナーを停止
            
            # 右クリックリスナーを開始
            self.mouse_listener = mouse.Listener(on_click=on_click)
            self.mouse_listener.start()
            
        except ImportError:
            # pynputが無い場合はフォールバック
            # pynputが無い場合はフォールバック
            self.check_for_right_click()
    
    def stop_global_click_detection(self):
        """グローバルクリック検出を停止"""
        try:
            if hasattr(self, 'mouse_listener'):
                self.mouse_listener.stop()
                delattr(self, 'mouse_listener')
        except Exception as e:
            logger.warning(f"クリック検出停止エラー: {e}")
    
    def on_right_click_detected(self, x, y):
        """右クリックが検出された時の処理"""
        self.position_display.configure(text=f"X: {x:4d} , Y: {y:4d}")
        self.x_entry.delete(0, "end")
        self.x_entry.insert(0, str(x))
        self.y_entry.delete(0, "end") 
        self.y_entry.insert(0, str(y))
        
        self.stop_tracking()
        
        # 確定音
        try:
            import winsound
            winsound.MessageBeep(winsound.MB_OK)
        except:
            pass
            
        messagebox.showinfo("座標確定", f"右クリックで座標が確定されました:\nX: {x}, Y: {y}", parent=self.dialog)
    
    def check_for_right_click(self):
        """フォールバック：定期的に右クリックをチェック（pynput無しの場合）"""
        if self.tracking:
            # 簡易的な実装：実際には完全ではないが参考用
            self.dialog.after(100, self.check_for_right_click)

    def close_dialog(self):
        """ダイアログを安全に閉じる"""
        self.stop_tracking()  # トラッキングとリスナーを停止
        self.dialog.destroy()
    
    def get_coordinates(self) -> Optional[Tuple[int, int]]:
        """選択された座標を取得"""
        self.dialog.wait_window()
        return self.selected_coordinates


class EnhancedImageDialog:
    """画像選択の拡張ダイアログ（クリップボード対応）"""
    
    def __init__(self, parent: tk.Tk, title: str = "画像選択"):
        self.parent = parent
        self.selected_path = None
        self.preview_label = None  # 明示的に初期化
        self.paste_area = None     # 明示的に初期化
        self.paste_label = None    # 明示的に初期化
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("700x600")  # クリック操作系で統一サイズに拡大
        self.dialog.configure(bg="#2b2b2b")
        self.dialog.resizable(True, True)
        self.dialog.minsize(500, 400)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # ウィンドウを中央に配置
        self.center_window()
        
        # UIを構築
        self.setup_ui()
        
        # ESCキーで閉じる
        self.dialog.bind('<Escape>', lambda e: self.dialog.destroy())
    
    def center_window(self):
        """ウィンドウをスクリーン中央に配置"""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f'+{x}+{y}')
    
    def setup_ui(self):
        """UIを構築（スクロール対応＋固定ボタン）"""
        # メインコンテナ
        container = tk.Frame(self.dialog, bg="#2b2b2b")
        container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # ボタンフレーム（最下部に固定）
        button_frame = tk.Frame(container, bg="#2b2b2b")
        button_frame.pack(side="bottom", fill="x", pady=(10, 0))
        
        # ttk.Buttonに変更して統一感を向上
        ok_btn = ttk.Button(
            button_frame,
            text="OK",
            command=self.confirm_selection,
            style='Primary.TButton',
            width=10
        )
        ok_btn.pack(side="right", padx=(10, 0))
        
        cancel_btn = ttk.Button(
            button_frame,
            text="キャンセル",
            command=self.dialog.destroy,
            style='Modern.TButton',
            width=10
        )
        cancel_btn.pack(side="right")
        
        # スクロール可能なメインコンテンツエリア
        canvas = tk.Canvas(container, bg="#2b2b2b", highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview, style='Modern.Vertical.TScrollbar')
        
        main_frame = tk.Frame(canvas, bg="#2b2b2b")
        
        # スクロール領域の設定
        def configure_scroll_region(event=None):
            canvas.update_idletasks()
            bbox = canvas.bbox("all")
            if bbox:
                # スクロール領域を正確に設定
                canvas.configure(scrollregion=bbox)
                
                # スクロールバーのつまみサイズを適切に設定
                content_height = bbox[3] - bbox[1]  # コンテンツの高さ
                canvas_height = canvas.winfo_height()  # 表示領域の高さ
                
                if content_height > canvas_height:
                    # 表示可能な比率を計算してスクロールバーに設定
                    visible_ratio = canvas_height / content_height
                    canvas.yview_moveto(0)  # 上端に移動
            else:
                canvas.configure(scrollregion=(0, 0, 0, 0))
        
        main_frame.bind("<Configure>", configure_scroll_region)
        
        # キャンバスウィンドウの設定
        canvas_window = canvas.create_window((0, 0), window=main_frame, anchor="nw")
        
        def configure_canvas_window(event):
            canvas.itemconfig(canvas_window, width=event.width)
            # キャンバスサイズ変更時にスクロール領域も再計算
            container.after_idle(configure_scroll_region)
        
        canvas.bind("<Configure>", configure_canvas_window)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # レイアウト配置
        canvas.pack(side="left", fill="both", expand=True, pady=(0, 10))
        
        # スクロールバーは必要な時だけ表示
        def check_scrollbar_visibility():
            canvas.update_idletasks()
            if canvas.bbox("all"):
                content_height = canvas.bbox("all")[3]
                canvas_height = canvas.winfo_height()
                if content_height > canvas_height:
                    scrollbar.pack(side="right", fill="y")
                else:
                    scrollbar.pack_forget()
        
        # 初期チェック
        container.after(100, check_scrollbar_visibility)
        
        # インスタンス変数として保存（後から呼び出せるように）
        self.check_scrollbar_visibility = check_scrollbar_visibility
        self.configure_scroll_region = configure_scroll_region
        
        # マウスホイールスクロール対応
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind("<MouseWheel>", on_mousewheel)  # Windows
        canvas.bind("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))  # Linux
        canvas.bind("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))   # Linux
        
        # タイトル
        title_label = tk.Label(
            main_frame,
            text="📷 画像を選択してください",
            font=("Meiryo UI", 14, "bold"),
            fg="#74c0fc",
            bg="#2b2b2b"
        )
        title_label.pack(pady=(0, 20))
        
        # 説明
        description_label = tk.Label(
            main_frame,
            text="📋 簡単な使い方:\n" +
                 "1️⃣ Shift+Win+S → 画面範囲選択 → Ctrl+V で貼り付け\n" +
                 "2️⃣ または下の「ファイルを選択」ボタンからファイル選択",
            font=("Meiryo UI", 10),
            fg="#ffffff",
            bg="#2b2b2b",
            justify="left"
        )
        description_label.pack(pady=(0, 15))
        
        # ファイル選択エリア
        file_frame = tk.Frame(main_frame, bg="#3c3c3c", relief="groove", bd=1)
        file_frame.pack(fill="x", pady=(0, 15))
        
        file_label = tk.Label(
            file_frame,
            text="📁 ファイルから選択:",
            font=("Meiryo UI", 11, "bold"),
            fg="#51cf66",
            bg="#3c3c3c"
        )
        file_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        file_btn = tk.Button(
            file_frame,
            text="📂 ファイルを選択...",
            command=self.select_file,
            font=("Meiryo UI", 10, "bold"),
            bg="#0d7377",
            fg="white",
            relief="flat",
            padx=20,
            pady=8,
            cursor="hand2"
        )
        file_btn.pack(pady=(0, 10))
        
        # クリップボード貼り付けエリア（拡大制限）
        clipboard_frame = tk.Frame(main_frame, bg="#3c3c3c", relief="groove", bd=1)
        clipboard_frame.pack(fill="x", pady=(0, 15))
        
        clipboard_label = tk.Label(
            clipboard_frame,
            text="📋 クリップボードから貼り付け:",
            font=("Meiryo UI", 11, "bold"),
            fg="#ffd93d",
            bg="#3c3c3c"
        )
        clipboard_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        # 貼り付けエリア（ドロップゾーン風）
        self.paste_area = tk.Frame(
            clipboard_frame,
            bg="#2b2b2b",
            relief="groove",  # dashedをgrooveに変更
            bd=2,
            height=120  # 高さを少し減らす
        )
        self.paste_area.pack(fill="x", padx=10, pady=(5, 10))
        
        self.paste_label = tk.Label(
            self.paste_area,
            text="🖼️ ここで Ctrl+V を押してください\n\n" +
                 "💡 Shift+Win+S で範囲選択した直後に\n" +
                 "   このダイアログで Ctrl+V を押すだけ！",
            font=("Meiryo UI", 11),
            fg="#adb5bd",
            bg="#2b2b2b"
        )
        self.paste_label.pack(expand=True)
        
        # プレビューエリア
        self.preview_label = None
        
        # キーバインド
        self.dialog.bind('<Control-v>', self.paste_from_clipboard)
        self.dialog.focus_set()
    
    def select_file(self):
        """ファイルを選択"""
        try:
            # ファイル選択開始
            
            # UIが完全に初期化されているかチェック
            if not hasattr(self, 'paste_area') or self.paste_area is None:
                # UIが初期化されていないため、ファイル選択を遅延実行
                self.dialog.after(100, self.select_file)  # 100ms後に再試行
                return
                
            file_path = filedialog.askopenfilename(
                title="画像ファイルを選択",
                filetypes=[("画像ファイル", "*.png *.jpg *.jpeg *.bmp *.gif *.tiff")],
                parent=self.dialog
            )
            if file_path:
                # ファイル選択成功
                self.selected_path = file_path
                self.show_preview(file_path, "ファイル")
        except Exception as e:
            logger.error(f"ファイル選択エラー: {e}")
            self.show_error_dialog("エラー", f"ファイル選択に失敗しました:\n{e}", parent=self.dialog)
    
    def paste_from_clipboard(self, event=None):
        """クリップボードから画像を貼り付け"""
        try:
            # クリップボード貼り付け開始
            
            # Pillowを使ってクリップボードから画像を取得
            from PIL import ImageGrab
            
            clipboard_image = ImageGrab.grabclipboard()
            # クリップボード画像取得
            
            if clipboard_image is None:
                # クリップボードに画像なし
                messagebox.showwarning(
                    "警告", 
                    "クリップボードに画像がありません。\n\n" +
                    "1. Shift+Win+S でスクリーンショットを撮影\n" +
                    "2. 範囲を選択\n" +
                    "3. このダイアログで Ctrl+V を押してください",
                    parent=self.dialog
                )
                return
            
            # 画像サイズ的取
            
            # 一時ファイルに保存
            temp_dir = Path("temp")
            temp_dir.mkdir(exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            temp_path = temp_dir / f"clipboard_image_{timestamp}.png"
            
            # 一時ファイルに保存
            clipboard_image.save(temp_path, "PNG")
            self.selected_path = str(temp_path)
            
            # プレビュー表示開始
            self.show_preview(str(temp_path), "クリップボード")
            
            # 貼り付け処理完了
            
        except ImportError as e:
            logger.error(f"PIL/Pillow ImportError: {e}")
            self.show_error_dialog("エラー", "PIL/Pillowが必要です。\npip install Pillow でインストールしてください。", parent=self.dialog)
        except Exception as e:
            logger.error(f"クリップボード貼り付けエラー: {e}")
            self.show_error_dialog("エラー", f"クリップボードからの貼り付けに失敗しました:\n{e}", parent=self.dialog)
    
    def show_preview(self, image_path: str, source: str):
        """画像プレビューを表示"""
        try:
            # プレビュー表示開始
            
            # paste_areaが存在するかチェック
            if not hasattr(self, 'paste_area') or self.paste_area is None:
                # paste_areaが初期化されていない
                self.show_error_dialog("エラー", "UIの初期化に失敗しました", parent=self.dialog)
                return
            
            # 既存のプレビューを削除
            if hasattr(self, 'preview_label') and self.preview_label:
                # 既存プレビューラベル削除
                self.preview_label.destroy()
            
            # 画像を読み込み
            pil_image = Image.open(image_path)
            # 画像読み込み完了
            
            # サムネイル作成
            pil_image.thumbnail((150, 80), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(pil_image)
            
            # プレビューラベル作成
            self.preview_label = tk.Label(
                self.paste_area,
                image=photo,
                bg="#2b2b2b"
            )
            self.preview_label.image = photo  # 参照を保持
            
            # 元のラベルを隠してプレビューを表示（拡張制限）
            if hasattr(self, 'paste_label') and self.paste_label:
                self.paste_label.pack_forget()
            self.preview_label.pack()
            
            # 情報表示
            file_size = os.path.getsize(image_path)
            size_str = f"{file_size:,} bytes" if file_size < 1024 else f"{file_size/1024:.1f} KB"
            
            info_text = f"✅ {source}画像が設定されました\n{os.path.basename(image_path)} ({size_str})"
            if hasattr(self, 'paste_label') and self.paste_label:
                self.paste_label.configure(text=info_text)
                self.paste_label.pack()
            
            # プレビュー表示完了
            
            # プレビュー表示後にスクロールバー可視性をチェック
            if hasattr(self, 'dialog'):
                self.dialog.after(200, self.check_scrollbar_visibility)
                # スクロール領域も再計算
                if hasattr(self, 'configure_scroll_region'):
                    self.dialog.after(250, self.configure_scroll_region)
            
        except Exception as e:
            logger.error(f"プレビュー表示エラー: {e}")
            self.show_error_dialog("エラー", f"画像プレビューの表示に失敗しました:\n{e}", parent=self.dialog)
    
    def confirm_selection(self):
        """選択を確定"""
        if self.selected_path:
            self.dialog.destroy()
        else:
            messagebox.showwarning(
                "警告", 
                "画像が選択されていません。\n\n" +
                "ファイルを選択するか、Shift+Win+S でスクリーンショットを撮って " +
                "Ctrl+V で貼り付けてください。",
                parent=self.dialog
            )
    
    def get_image_path(self) -> Optional[str]:
        """選択された画像パスを取得"""
        self.dialog.wait_window()
        return self.selected_path



class ConfigSwitcherDialog:
    """設定ファイル切替ダイアログ"""
    def __init__(self, parent: tk.Tk, json_files: List[Dict], load_callback):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("⚙️ 設定切替")
        self.dialog.geometry("700x600")
        self.dialog.resizable(True, True)
        self.dialog.minsize(500, 400)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # ダイアログを中央に配置
        self.center_window()
        
        # スタイル設定
        self.dialog.configure(bg=AppConfig.THEME['bg_primary'])
        
        self.json_files = json_files
        self.load_callback = load_callback
        self.selected_file = None
        self.setup_ui()
        
        # ESCキーで閉じる
        self.dialog.bind('<Escape>', lambda e: self.dialog.destroy())
    
    def center_window(self):
        """ウィンドウをスクリーン中央に配置"""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f'+{x}+{y}')
    
    def setup_ui(self):
        """UIを構築"""
        # メインフレーム
        main_frame = ttk.Frame(self.dialog)
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # タイトルラベル
        title_label = ttk.Label(main_frame, text="設定ファイルを選択してください", 
                               font=('Segoe UI', 12, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # ファイル一覧フレーム
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill='both', expand=True, pady=(0, 15))
        
        # リストボックスとスクロールバー
        list_container = ttk.Frame(list_frame)
        list_container.pack(fill='both', expand=True)
        
        # リストボックス
        self.listbox = tk.Listbox(list_container, 
                                 font=('Segoe UI', 10),
                                 selectmode='single',
                                 height=15)
        self.listbox.pack(side='left', fill='both', expand=True)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        
        # ファイル一覧を追加
        for file_info in self.json_files:
            self.listbox.insert(tk.END, file_info["display"])
        
        # ダブルクリックで読み込み
        self.listbox.bind('<Double-Button-1>', self.on_double_click)
        self.listbox.bind('<<ListboxSelect>>', self.on_select)
        
        # プレビューエリア
        preview_frame = ttk.LabelFrame(main_frame, text="ファイル情報")
        preview_frame.pack(fill='x', pady=(0, 15))
        
        self.info_text = tk.Text(preview_frame, height=3, font=('Consolas', 9), 
                                wrap='word', state='disabled')
        self.info_text.pack(fill='x', padx=10, pady=5)
        
        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x')
        
        # ボタン（統一スタイル適用）
        ttk.Button(button_frame, text="OK", command=self.load_selected, style='Primary.TButton', width=10).pack(side='right', padx=(10, 0))
        ttk.Button(button_frame, text="キャンセル", command=self.dialog.destroy, style='Modern.TButton', width=10).pack(side='right')
        ttk.Button(button_frame, text="フォルダを開く", command=self.open_config_folder, style='Modern.TButton', width=12).pack(side='left')
        
        # 初期選択
        if self.json_files:
            self.listbox.select_set(0)
            self.on_select(None)
    
    def on_select(self, event):
        """リスト選択時の処理"""
        selection = self.listbox.curselection()
        if selection:
            index = selection[0]
            self.selected_file = self.json_files[index]
            self.update_preview()
    
    def on_double_click(self, event):
        """ダブルクリック時の処理"""
        self.load_selected()
    
    def update_preview(self):
        """プレビューエリアを更新"""
        if not self.selected_file:
            return
        
        self.info_text.config(state='normal')
        self.info_text.delete('1.0', tk.END)
        
        info_lines = [
            f"ファイル名: {self.selected_file['name']}",
            f"ステップ数: {self.selected_file['steps']}",
            f"ファイルサイズ: {self.selected_file['size']} bytes",
            f"パス: {self.selected_file['path']}"
        ]
        
        self.info_text.insert('1.0', '\n'.join(info_lines))
        self.info_text.config(state='disabled')
    
    def load_selected(self):
        """選択されたファイルを読み込み"""
        if not self.selected_file:
            return
        
        try:
            self.load_callback(self.selected_file['path'])
            self.dialog.destroy()
        except Exception as e:
            import tkinter.messagebox as messagebox
            messagebox.showerror("エラー", f"設定の読み込みに失敗しました:\n{e}", parent=self.dialog)
    
    def open_config_folder(self):
        """configフォルダを開く"""
        try:
            import subprocess
            config_dir = os.path.dirname(self.json_files[0]['path']) if self.json_files else os.path.join(os.path.dirname(__file__), "config")
            subprocess.run(['explorer', config_dir], check=True)
        except Exception as e:
            import tkinter.messagebox as messagebox
            messagebox.showerror("エラー", f"フォルダを開けませんでした:\n{e}", parent=self.dialog)


if __name__ == "__main__":
    try:
        print("=== Auto GUI Tool Professional v2.0 ===")
        print("アプリケーションを起動しています...")
        
        root = tk.Tk()
        app = AutoActionTool(root)
        
        print("GUIウィンドウを表示しています...")
        print("アプリケーションが正常に起動しました。")
        print("ウィンドウを閉じるまでこのターミナルは開いたままにしてください。")
        
        root.mainloop()
        
        print("アプリケーションが終了しました。")
        
    except Exception as e:
        print(f"エラー: アプリケーションの起動に失敗しました: {e}")
        logger.error(f"アプリケーション起動エラー: {e}")
        try:
            messagebox.showerror("エラー", f"アプリケーションの起動に失敗しました: {e}")
        except:
            pass  # GUIが利用できない場合はスキップ
