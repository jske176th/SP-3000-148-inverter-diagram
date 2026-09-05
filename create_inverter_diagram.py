#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SP-3000-148 DC48V → AC100V 3000W インバータ ブロック図
PowerPoint生成スクリプト
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

def add_title_slide(prs, title, subtitle):
    """タイトルスライドを追加"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白レイアウト
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    # タイトル
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_frame.word_wrap = True
    title_p = title_frame.paragraphs[0]
    title_p.font.size = Pt(54)
    title_p.font.bold = True
    title_p.font.color.rgb = RGBColor(0, 0, 0)
    title_p.alignment = PP_ALIGN.CENTER
    
    # サブタイトル
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(3.8), Inches(9), Inches(1))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = subtitle
    subtitle_p = subtitle_frame.paragraphs[0]
    subtitle_p.font.size = Pt(28)
    subtitle_p.font.color.rgb = RGBColor(64, 64, 64)
    subtitle_p.alignment = PP_ALIGN.CENTER

def add_text_box(slide, left, top, width, height, text, font_size=11, bold=False, alignment=PP_ALIGN.LEFT):
    """テキストボックスを追加"""
    textbox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    text_frame = textbox.text_frame
    text_frame.text = text
    text_frame.word_wrap = True
    p = text_frame.paragraphs[0]
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = alignment
    return textbox

def add_rectangle(slide, left, top, width, height, text, fill_color=(240, 240, 240), text_color=(0, 0, 0)):
    """矩形ボックスを追加"""
    shape = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*fill_color)
    shape.line.color.rgb = RGBColor(0, 0, 0)
    shape.line.width = Pt(1.5)
    
    text_frame = shape.text_frame
    text_frame.text = text
    text_frame.word_wrap = True
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    
    for paragraph in text_frame.paragraphs:
        paragraph.font.size = Pt(10)
        paragraph.font.bold = True
        paragraph.font.color.rgb = RGBColor(*text_color)
        paragraph.alignment = PP_ALIGN.CENTER
    
    return shape

def add_arrow(slide, x1, y1, x2, y2, color=(0, 0, 0), width=2):
    """矢印を追加"""
    connector = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    connector.line.color.rgb = RGBColor(*color)
    connector.line.width = Pt(width)

def create_block_diagram_slide(prs):
    """ブロック図スライドを作成"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    # タイトル
    add_text_box(slide, 0.5, 0.3, 9, 0.5, "全体システムブロック図", font_size=24, bold=True, alignment=PP_ALIGN.CENTER)
    
    # ===== 上段 =====
    # フロントパネル
    add_rectangle(slide, 3.5, 1.2, 2.5, 0.6, "フロントパネル\n電源SW / LED表示", fill_color=(200, 200, 200))
    
    # 矢印
    add_arrow(slide, 4.75, 1.8, 4.75, 2.2, width=2)
    
    # CN3(16P)コネクタ
    add_text_box(slide, 3.8, 2.0, 1.9, 0.3, "CN3(16P)", font_size=9)
    
    # 上段制御基板
    add_rectangle(slide, 3.2, 2.4, 3.1, 0.8, "上段制御基板\n66-BA97-2101\nMCU / DIP=010 / ブザー", fill_color=(220, 220, 220))
    
    # 矢印
    add_arrow(slide, 4.75, 3.2, 4.75, 3.6, width=2)
    
    # 上段電力基板
    add_rectangle(slide, 2.8, 3.8, 3.9, 1.2, 
                  "上段電力基板\nDC/DC昇圧 / 高周波トランス\nHブリッジPWM / LCフィルタ", 
                  fill_color=(240, 240, 240))
    
    # DC48V入力（左）
    add_arrow(slide, 0.5, 4.4, 2.8, 4.4, width=2.5)
    add_text_box(slide, 0.8, 4.15, 1.5, 0.25, "DC48V入力", font_size=10, bold=True)
    
    # AC100V出力（右）
    add_arrow(slide, 6.7, 4.4, 8.5, 4.4, width=2.5)
    add_text_box(slide, 7.3, 4.15, 1.5, 0.25, "AC100V出力", font_size=10, bold=True)
    
    # ===== 下段 =====
    # 矢印（ケーブル）
    add_arrow(slide, 4.75, 5.0, 4.75, 5.4, width=2)
    add_text_box(slide, 3.5, 5.1, 2.5, 0.25, "ケーブル接続", font_size=9)
    
    # リアパネル基板
    add_rectangle(slide, 3.5, 5.6, 2.5, 0.5, "リアパネル基板", fill_color=(200, 200, 200))
    
    # 矢印
    add_arrow(slide, 4.75, 6.1, 4.75, 6.4, width=2)
    
    # 外部リモートコネクタ
    add_rectangle(slide, 3.2, 6.6, 3.1, 0.5, "外部リモートコネクタ", fill_color=(200, 200, 200))
    
    # ===== 左側の下段ブロック =====
    # 下段制御基板
    add_rectangle(slide, 0.3, 3.8, 2.3, 0.8, "下段制御基板\n66-BA97-2101\nMCU / DIP=001", fill_color=(220, 220, 220))
    
    # 矢印
    add_arrow(slide, 1.45, 4.6, 1.45, 5.0, width=2)
    
    # 下段電力基板
    add_rectangle(slide, 0.05, 5.2, 2.9, 1.2, 
                  "下段電力基板\nDC/DC昇圧 / 高周波トランス\nHブリッジPWM / LCフィルタ", 
                  fill_color=(240, 240, 240))
    
    # DC48V入力（左）
    add_arrow(slide, 0.05, 5.8, 0.05, 5.8, width=2.5)
    
    # CN2通信ライン
    add_arrow(slide, 2.55, 4.2, 3.2, 4.2, width=2)
    add_text_box(slide, 2.5, 3.95, 1.2, 0.25, "CN2(8P)通信", font_size=8, bold=True)

def create_control_flow_slide(prs):
    """制御フロー説明スライドを作成"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    # タイトル
    add_text_box(slide, 0.5, 0.3, 9, 0.5, "制御構成と通信フロー", font_size=24, bold=True, alignment=PP_ALIGN.CENTER)
    
    # 左側：上段システム
    add_text_box(slide, 0.5, 1.1, 1.2, 0.3, "【上段システム】", font_size=12, bold=True)
    add_rectangle(slide, 0.3, 1.5, 2.2, 0.7, "上段制御基板\n(マスター)\nMCU DIP=010", fill_color=(220, 220, 220))
    add_arrow(slide, 1.4, 2.2, 1.4, 2.5, width=2)
    add_rectangle(slide, 0.3, 2.5, 2.2, 0.7, "上段電力基板\nPWM制御\n昇圧・逆変換", fill_color=(240, 240, 240))
    
    # 中央：同期通信
    add_text_box(slide, 2.8, 1.7, 1.9, 0.3, "CN2(8P)通信", font_size=11, bold=True, alignment=PP_ALIGN.CENTER)
    add_arrow(slide, 2.6, 1.95, 4.0, 1.95, width=2.5)
    add_text_box(slide, 2.6, 2.15, 1.9, 0.4, "・同期信号\n・電流共有\n・異常通知", font_size=9)
    add_arrow(slide, 4.0, 1.95, 5.4, 1.95, width=2.5)
    
    # 右側：下段システム
    add_text_box(slide, 6.0, 1.1, 1.2, 0.3, "【下段システム】", font_size=12, bold=True)
    add_rectangle(slide, 5.8, 1.5, 2.2, 0.7, "下段制御基板\n(スレーブ)\nMCU DIP=001", fill_color=(220, 220, 220))
    add_arrow(slide, 6.9, 2.2, 6.9, 2.5, width=2)
    add_rectangle(slide, 5.8, 2.5, 2.2, 0.7, "下段電力基板\nPWM制御\n昇圧・逆変換", fill_color=(240, 240, 240))
    
    # 詳細説明
    add_text_box(slide, 0.5, 3.5, 9, 3.0, 
                 "■ 制御フロー詳細\n\n"
                 "1. 電源制御：フロントパネル電源SWで上下段の電源を同時ON/OFF\n\n"
                 "2. 入力電力分配：DC48V入力を上下段に分配し、各段で独立した昇圧・逆変換を実行\n\n"
                 "3. 同期制御：CN2通信ラインにより、上段（マスター）が下段（スレーブ）と同期\n"
                 "   - PWM周波数・位相の同期\n"
                 "   - 出力電流の均等分配（並列動作）\n"
                 "   - 異常検出時の同時シャットダウン\n\n"
                 "4. 出力合成：両段のAC100V出力を並列して最大3000Wの電力供給を実現\n\n"
                 "5. 保護機能：過電流、過温度、低入力電圧時にブザーで異常を通知",
                 font_size=10)

def create_power_flow_slide(prs):
    """電力フロー説明スライドを作成"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    # タイトル
    add_text_box(slide, 0.5, 0.3, 9, 0.5, "電力フロー構成", font_size=24, bold=True, alignment=PP_ALIGN.CENTER)
    
    # 左側：上段電力フロー
    add_text_box(slide, 0.3, 1.1, 2.0, 0.3, "【上段：電力フロー】", font_size=12, bold=True)
    add_arrow(slide, 0.5, 1.6, 1.5, 1.6, width=2.5)
    add_text_box(slide, 0.3, 1.4, 1.8, 0.25, "DC48V入力", font_size=10, bold=True)
    
    add_rectangle(slide, 1.5, 1.4, 1.3, 0.4, "DC/DC昇圧", fill_color=(230, 230, 230))
    add_arrow(slide, 2.8, 1.6, 3.5, 1.6, width=2)
    add_rectangle(slide, 3.5, 1.4, 1.3, 0.4, "高周波\nトランス", fill_color=(230, 230, 230))
    
    add_arrow(slide, 4.8, 1.6, 5.3, 1.6, width=2)
    add_rectangle(slide, 5.3, 1.4, 1.0, 0.4, "Hブリッジ\nPWM", fill_color=(230, 230, 230))
    
    add_arrow(slide, 6.3, 1.6, 7.0, 1.6, width=2)
    add_rectangle(slide, 7.0, 1.4, 1.2, 0.4, "LCフィルタ", fill_color=(230, 230, 230))
    
    add_arrow(slide, 8.2, 1.6, 8.8, 1.6, width=2.5)
    add_text_box(slide, 8.5, 1.4, 1.2, 0.25, "AC100V出力\n最大1500W", font_size=9, bold=True)
    
    # 右側：下段電力フロー
    add_text_box(slide, 0.3, 2.3, 2.0, 0.3, "【下段：電力フロー】", font_size=12, bold=True)
    add_arrow(slide, 0.5, 2.8, 1.5, 2.8, width=2.5)
    add_text_box(slide, 0.3, 2.6, 1.8, 0.25, "DC48V入力", font_size=10, bold=True)
    
    add_rectangle(slide, 1.5, 2.6, 1.3, 0.4, "DC/DC昇圧", fill_color=(230, 230, 230))
    add_arrow(slide, 2.8, 2.8, 3.5, 2.8, width=2)
    add_rectangle(slide, 3.5, 2.6, 1.3, 0.4, "高周波\nトランス", fill_color=(230, 230, 230))
    
    add_arrow(slide, 4.8, 2.8, 5.3, 2.8, width=2)
    add_rectangle(slide, 5.3, 2.6, 1.0, 0.4, "Hブリッジ\nPWM", fill_color=(230, 230, 230))
    
    add_arrow(slide, 6.3, 2.8, 7.0, 2.8, width=2)
    add_rectangle(slide, 7.0, 2.6, 1.2, 0.4, "LCフィルタ", fill_color=(230, 230, 230))
    
    add_arrow(slide, 8.2, 2.8, 8.8, 2.8, width=2.5)
    add_text_box(slide, 8.5, 2.6, 1.2, 0.25, "AC100V出力\n最大1500W", font_size=9, bold=True)
    
    # 並列合成
    add_text_box(slide, 3.5, 3.5, 3.0, 0.3, "【出力並列合成】", font_size=12, bold=True, alignment=PP_ALIGN.CENTER)
    add_arrow(slide, 5.0, 2.0, 5.0, 3.8, width=2)
    add_arrow(slide, 5.0, 2.8, 5.0, 3.8, width=2)
    add_rectangle(slide, 4.0, 4.0, 2.0, 0.5, "AC並列接続\n最大3000W", fill_color=(200, 200, 200))
    
    # 詳細説明
    add_text_box(slide, 0.5, 4.7, 9, 2.0, 
                 "■ 電力処理ステップ\n\n"
                 "1. DC/DC昇圧：入力DC48Vを昇圧してHブリッジ駆動用DC電圧を生成（効率：～95%）\n\n"
                 "2. 高周波トランス：昇圧DC電圧を絶縁トランスで変換（高周波で小型化）\n\n"
                 "3. Hブリッジ逆変換：MCUのPWM制御信号でDC→AC逆変換（周波数：50/60Hz選択可能）\n\n"
                 "4. LCフィルタ：PWMのスイッチング周波数を除去して純正弦波AC100V出力を生成\n\n"
                 "5. 並列出力：上下段の出力を並列して3000Wの合計出力を実現（負荷分散）",
                 font_size=9)

def create_interface_slide(prs):
    """インターフェース説明スライドを作成"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    # タイトル
    add_text_box(slide, 0.5, 0.3, 9, 0.5, "インターフェース仕様", font_size=24, bold=True, alignment=PP_ALIGN.CENTER)
    
    # CN3（フロントパネル接続）
    add_text_box(slide, 0.5, 1.1, 2.5, 0.3, "■ CN3(16P) - フロントパネル", font_size=11, bold=True)
    add_rectangle(slide, 0.3, 1.5, 2.8, 1.8, 
                  "用途：フロントパネルとの接続\n\n"
                  "信号構成：\n"
                  "・電源SW（入力）\n"
                  "・LED表示信号（出力）\n"
                  "・ブザー制御（出力）\n"
                  "・ステータス信号（出力）\n\n"
                  "コネクタ：16ピン DIN型",
                  fill_color=(240, 240, 240), text_color=(0, 0, 0))
    
    # CN2（制御基板間通信）
    add_text_box(slide, 3.5, 1.1, 2.5, 0.3, "■ CN2(8P) - 制御基板間通信", font_size=11, bold=True)
    add_rectangle(slide, 3.3, 1.5, 2.8, 1.8, 
                  "用途：上下段MCU間の通信\n\n"
                  "通信内容：\n"
                  "・PWM同期信号\n"
                  "・出力電流情報\n"
                  "・異常通知\n"
                  "・スレーブ制御\n\n"
                  "コネクタ：8ピン DIN型",
                  fill_color=(240, 240, 240), text_color=(0, 0, 0))
    
    # 外部リモート
    add_text_box(slide, 6.5, 1.1, 2.5, 0.3, "■ 外部リモートコネクタ", font_size=11, bold=True)
    add_rectangle(slide, 6.3, 1.5, 2.8, 1.8, 
                  "用途：外部制御入力\n\n"
                  "機能：\n"
                  "・リモート電源制御\n"
                  "・状態監視（出力可能）\n"
                  "・異常信号出力\n\n"
                  "リアパネル基板経由で\n"
                  "制御基板と接続",
                  fill_color=(240, 240, 240), text_color=(0, 0, 0))
    
    # 制御基板仕様
    add_text_box(slide, 0.5, 3.6, 9, 0.35, "■ 制御基板仕様（66-BA97-2101）", font_size=11, bold=True)
    add_rectangle(slide, 0.3, 4.1, 4.3, 2.3, 
                  "【上段制御基板】\nDIP設定：0 1 0\n（マスター・同期トリガー・通常モード）\n\n"
                  "機能：\n"
                  "・フロントパネル制��\n"
                  "・PWM波形生成\n"
                  "・電流検出＆フィードバック\n"
                  "・異常検出＆通知\n"
                  "・ブザー制御",
                  fill_color=(220, 220, 220), text_color=(0, 0, 0))
    
    add_rectangle(slide, 4.7, 4.1, 4.3, 2.3, 
                  "【下段制御基板】\nDIP設定：0 0 1\n（スレーブ・同期受信・通常モード）\n\n"
                  "機能：\n"
                  "・上段との同期受信\n"
                  "・PWM波形生成（同期）\n"
                  "・電流検出＆フィードバック\n"
                  "・異常検出＆通知",
                  fill_color=(220, 220, 220), text_color=(0, 0, 0))

def create_specifications_slide(prs):
    """仕様詳細スライドを作成"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    # タイトル
    add_text_box(slide, 0.5, 0.3, 9, 0.5, "システム仕様", font_size=24, bold=True, alignment=PP_ALIGN.CENTER)
    
    # 入出力仕様
    add_text_box(slide, 0.5, 1.1, 4.3, 0.3, "■ 入出力仕様", font_size=11, bold=True)
    add_rectangle(slide, 0.3, 1.5, 4.5, 2.0, 
                  "【入力】\n"
                  "・電圧：DC48V ±10%\n"
                  "・最大入力電流：上下段各65A（合計130A）\n"
                  "・最大入力電力：6240W（DC48V×130A）\n\n"
                  "【出力】\n"
                  "・電圧：AC100V ±5%\n"
                  "・周波数：50/60Hz（選択）\n"
                  "・最大出力電力：3000W（上下段合計1500W×2）\n"
                  "・波形：純正弦波（THD≦3%）",
                  fill_color=(240, 240, 240), text_color=(0, 0, 0))
    
    # パフォーマンス
    add_text_box(slide, 5.2, 1.1, 4.3, 0.3, "■ パフォーマンス", font_size=11, bold=True)
    add_rectangle(slide, 5.0, 1.5, 4.5, 2.0, 
                  "■ 変換効率：≧92%（推定）\n"
                  "・DC/DC昇圧段：～95%\n"
                  "・Hブリッジ逆変換：～97%\n"
                  "・トランス＆フィルタ損失：～98%\n\n"
                  "■ 応答速度\n"
                  "・起動時間：＜500ms\n"
                  "・同期ロック時間：＜100ms\n"
                  "・異常検出：リアルタイム\n\n"
                  "■ 冷却：強制空冷または自然対流",
                  fill_color=(240, 240, 240), text_color=(0, 0, 0))
    
    # 動作仕様
    add_text_box(slide, 0.5, 3.7, 9, 0.3, "■ 動作仕様", font_size=11, bold=True)
    add_rectangle(slide, 0.3, 4.1, 9.2, 1.8, 
                  "■ 動作環境\n"
                  "・温度範囲：-10℃～+50℃（推奨0～40℃）  ・湿度範囲：20～80%（結露なし）  ・高度：＜2000m\n\n"
                  "■ 保護機能\n"
                  "・過電流保護：全出力ラインに実装（自動シャットダウン）\n"
                  "・過温度保護：温度センサで監視（警告→シャットダウン）\n"
                  "・低入力電圧保護：DC48V＜43.2V時に遮断\n"
                  "・同期エラー保護：CN2通信断時に下段自動OFF\n"
                  "・異常通知：ブザー＆LED表示で即座に表示\n\n"
                  "■ 絶縁耐圧\n"
                  "・入力～出力間：DC3000V, 1分間（安全基準対応）\n"
                  "・信号回路間：DC500V, 1分間",
                  fill_color=(240, 240, 240), text_color=(0, 0, 0))

def main():
    """メイン処理"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    print("PowerPoint資料を生成中...")
    
    # スライド1：タイトル
    print("  - スライド1：タイトルページ")
    add_title_slide(prs, "SP-3000-148", "DC48V → AC100V 3000W インバータ\nシステムブロック図詳細")
    
    # スライド2：全体ブロック図
    print("  - スライド2：全体システムブロック図")
    create_block_diagram_slide(prs)
    
    # スライド3：制御フロー
    print("  - スライド3：制御構成と通信フロー")
    create_control_flow_slide(prs)
    
    # スライド4：電力フロー
    print("  - スライド4：電力フロー構成")
    create_power_flow_slide(prs)
    
    # スライド5：インターフェース
    print("  - スライド5：インターフェース仕様")
    create_interface_slide(prs)
    
    # スライド6：仕様詳細
    print("  - スライド6：システム仕様")
    create_specifications_slide(prs)
    
    # 保存
    output_file = "SP-3000-148_Inverter_BlockDiagram.pptx"
    prs.save(output_file)
    print(f"\n✓ PowerPoint資料を保存しました: {output_file}")
    print(f"  - スライド数：6枚")
    print(f"  - サイズ：10 x 7.5 インチ（横置き）")
    print(f"  - カラースキーム：白黒")

if __name__ == "__main__":
    main()
