#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SP-3000-148 PowerPoint 日本語→英語変換スクリプト
SP-3000-148_Inverter_BlockDiagram2.pptx を英語版に変換
"""

from pptx import Presentation
from pptx.util import Inches, Pt

# 日本語→英語の翻訳辞書
translation_dict = {
    # === ブロック図用語 ===
    "フロントパネル": "Front Panel",
    "電源SW": "Power Switch",
    "LED表示": "LED Display",
    "CN3(16P)": "CN3(16P)",
    "上段制御基板": "Upper Stage Control Board",
    "66-BA97-2101": "66-BA97-2101",
    "MCU": "MCU",
    "DIP=0 1 0": "DIP=0 1 0",
    "DIP=010": "DIP=010",
    "ブザー": "Buzzer",
    "電力制御I/F": "Power Control I/F",
    "上段電力基板": "Upper Stage Power Board",
    "DC/DC昇圧": "DC/DC Boost",
    "高周波トランス": "High Frequency Transformer",
    "HブリッジPWM": "H-Bridge PWM",
    "LCフィルタ": "LC Filter",
    "DC48V入力": "DC48V Input",
    "AC100V出力": "AC100V Output",
    "ケーブル": "Cable",
    "ケーブル接続": "Cable Connection",
    "リアパネル基板": "Rear Panel Board",
    "外部リモートコネクタ": "External Remote Connector",
    "下段制御基板": "Lower Stage Control Board",
    "DIP=0 0 1": "DIP=0 0 1",
    "DIP=001": "DIP=001",
    "下段電力基板": "Lower Stage Power Board",
    
    # === スライドタイトル ===
    "全体システムブロック図": "Overall System Block Diagram",
    "制御構成と通信フロー": "Control Configuration and Communication Flow",
    "電力フロー構成": "Power Flow Configuration",
    "インターフェース仕様": "Interface Specifications",
    "システム仕様": "System Specifications",
    
    # === 制御フロー用語 ===
    "【上段システム】": "[Upper Stage System]",
    "【下段システム】": "[Lower Stage System]",
    "CN2(8P)通信": "CN2(8P) Communication",
    "マスター": "Master",
    "スレーブ": "Slave",
    "PWM制御": "PWM Control",
    "昇圧・逆変換": "Boost & Inversion",
    
    # === 通信内容 ===
    "・同期信号": "• Sync Signal",
    "・電流共有": "• Current Sharing",
    "・異常通知": "• Fault Notification",
    
    # === 電力フロー詳細 ===
    "【電力フロー】": "[Power Flow]",
    "【出力並列合成】": "[Output Parallel Combination]",
    "AC並列接続": "AC Parallel Connection",
    "最大3000W": "Max 3000W",
    "最大1500W": "Max 1500W",
    
    # === インターフェース ===
    "■ CN3(16P) - フロントパネル": "■ CN3(16P) - Front Panel",
    "■ CN2(8P) - 制御基板間通信": "■ CN2(8P) - Inter-Board Communication",
    "■ 外部リモートコネクタ": "■ External Remote Connector",
    "■ 制御基板仕様（66-BA97-2101）": "■ Control Board Specifications (66-BA97-2101)",
    "【上段制御基板】": "[Upper Stage Control Board]",
    "【下段制御基板】": "[Lower Stage Control Board]",
    "DIP設定": "DIP Setting",
    "通常モード": "Normal Mode",
    "用途": "Purpose",
    "信号構成": "Signal Configuration",
    "コネクタ": "Connector",
    "通信内容": "Communication Content",
    "機能": "Function",
    "ピン DIN型": "Pin DIN Type",
    
    # === 機能説明 ===
    "・フロントパネル制御": "• Front Panel Control",
    "・PWM波形生成": "• PWM Waveform Generation",
    "・電流検出＆フィードバック": "• Current Detection & Feedback",
    "・異常検出＆通知": "• Fault Detection & Notification",
    "・ブザー制御": "• Buzzer Control",
    "・上段との同期受信": "• Synchronization with Upper Stage",
    
    # === リモート機能 ===
    "・リモート電源制御": "• Remote Power Control",
    "・状態監視（出力可能）": "• Status Monitoring (Output available)",
    "・異常信号出力": "• Fault Signal Output",
    "リアパネル基板経由で": "via Rear Panel Board",
    "制御基板と接続": "connected to Control Board",
    
    # === 制御フロー詳細 ===
    "■ 制御フロー詳細": "■ Control Flow Details",
    "1. 電源制御": "1. Power Control",
    "2. 入力電力分配": "2. Input Power Distribution",
    "3. 同期制御": "3. Synchronization Control",
    "4. 出力合成": "4. Output Synthesis",
    "5. 保護機能": "5. Protection Features",
    
    # === 電力処理ステップ ===
    "■ 電力処理ステップ": "■ Power Processing Steps",
    "1. DC/DC昇圧": "1. DC/DC Boost",
    "2. 高周波トランス": "2. High Frequency Transformer",
    "3. Hブリッジ逆変換": "3. H-Bridge Inversion",
    "4. LCフィルタ": "4. LC Filter",
    "5. 並列出力": "5. Parallel Output",
    
    # === 仕様詳細 ===
    "■ 入出力仕様": "■ Input/Output Specifications",
    "■ パフォーマンス": "■ Performance",
    "■ 動作仕様": "■ Operating Specifications",
    
    "【入力】": "[Input]",
    "【出力】": "[Output]",
    
    "・電圧": "• Voltage",
    "・最大入力電流": "• Maximum Input Current",
    "・最大入力電力": "• Maximum Input Power",
    "・周波数": "• Frequency",
    "・最大出力電力": "• Maximum Output Power",
    "・波形": "• Waveform",
    
    "■ 変換効率": "■ Conversion Efficiency",
    "■ 応答速度": "■ Response Speed",
    "■ 冷却": "■ Cooling",
    
    "DC48V": "DC48V",
    "AC100V": "AC100V",
    "±10%": "±10%",
    "±5%": "±5%",
    "50/60Hz（選択）": "50/60Hz (Selectable)",
    "純正弦波（THD≦3%）": "Pure Sine Wave (THD≦3%)",
    "≧92%（推定）": "≧92% (Estimated)",
    "～95%": "~95%",
    "～97%": "~97%",
    "～98%": "~98%",
    
    "・起動時間": "• Startup Time",
    "・同期ロック時間": "• Sync Lock Time",
    "・異常検出": "• Fault Detection",
    "強制空冷または自然対流": "Forced Air Cooling or Natural Convection",
    
    "■ 動作環境": "■ Operating Environment",
    "・温度範囲": "• Temperature Range",
    "・湿度範囲": "• Humidity Range",
    "・高度": "• Altitude",
    "結露なし": "No Condensation",
    
    "■ 保護機能": "■ Protection Features",
    "・過電流保護": "• Overcurrent Protection",
    "・過温度保護": "• Overtemperature Protection",
    "・低入力電圧保護": "• Low Input Voltage Protection",
    "・同期エラー保護": "• Synchronization Error Protection",
    "・異常通知": "• Fault Notification",
    
    "■ 絶縁耐圧": "■ Dielectric Withstand Voltage",
    "・入力～出力間": "• Input to Output",
    "・信号回路間": "• Between Signal Circuits",
    
    # === 詳細説明テキスト ===
    "フロントパネル電源SWで上下段の電源を同時ON/OFF": "Front panel power switch simultaneously turns upper and lower stages ON/OFF",
    "DC48V入力を上下段に分配し、各段で独立した昇圧・逆変換を実行": "DC48V input is distributed to upper and lower stages for independent boost and inversion",
    "PWM周波数・位相の同期": "PWM frequency and phase synchronization",
    "出力電流の均等分配（並列動作）": "Equal distribution of output current (parallel operation)",
    "異常検出時の同時シャットダウン": "Simultaneous shutdown on fault detection",
    "両段のAC100V出力を並列して最大3000Wの電力供給を実現": "Parallel AC100V outputs from both stages achieve 3000W maximum power supply",
    "過電流、過温度、低入力電圧時にブザーで異常を通知": "Buzzer alerts for overcurrent, overtemperature, and low input voltage",
    
    "昇圧DC電圧を絶縁トランスで変換（高周波で小型化）": "Boost DC voltage converted by isolation transformer (compact with high frequency)",
    "MCUのPWM制御信号でDC→AC逆変換（周波数：50/60Hz選択可能）": "DC to AC inversion via MCU PWM control (Selectable 50/60Hz)",
    "PWMのスイッチング周波数を除去して純正弦波AC100V出力を生成": "PWM switching frequency removed to generate pure sine wave AC100V output",
    "上下段の出力を並列して3000Wの合計出力を実現（負荷分散）": "Parallel outputs from both stages achieve 3000W total (load distribution)",
    
    "上段（マスター）が下段（スレーブ）と同期": "Upper stage (Master) synchronizes with lower stage (Slave)",
    "・起動時間：＜500ms": "• Startup time: <500ms",
    "・同期ロック時間：＜100ms": "• Sync lock time: <100ms",
    "・異常検出：リアルタイム": "• Fault detection: Real-time",
    
    "効率：～95%": "Efficiency: ~95%",
    "効率：～97%": "Efficiency: ~97%",
    "効率：～98%": "Efficiency: ~98%",
    "自動シャットダウン": "Auto Shutdown",
    "温度センサで監視（警告→シャットダウン）": "Temperature sensor monitoring (Warning → Shutdown)",
    "DC48V＜43.2V時に遮断": "Cutoff when DC48V <43.2V",
    "CN2通信断時に下段自動OFF": "Lower stage automatically OFF on CN2 communication failure",
    "ブザー＆LED表示で即座に表示": "Buzzer & LED display immediate indication",
    "DC3000V, 1分間（安全基準対応）": "DC3000V, 1 minute (Safety standard compliant)",
    "DC500V, 1分間": "DC500V, 1 minute",
    
    "-10℃～+50℃（推奨0～40℃）": "-10°C to +50°C (Recommended 0-40°C)",
    "20～80%（結露なし）": "20-80% (No condensation)",
    "＜2000m": "<2000m",
    
    "マスター・同期トリガー・通常モード": "Master, Sync Trigger, Normal Mode",
    "スレーブ・同期受信・通常モード": "Slave, Sync Reception, Normal Mode",
    
    "上下段MCU間の通信": "Communication between upper and lower MCUs",
    "外部制御入力": "External Control Input",
    "ピン": "Pin",
    "型": "Type",
}

def translate_presentation(input_path, output_path):
    """PowerPointプレゼンテーション全体を翻訳"""
    print(f"ファイルを読み込み中: {input_path}")
    prs = Presentation(input_path)
    
    translated_count = 0
    
    # スライドごとに処理
    for slide_idx, slide in enumerate(prs.slides, 1):
        print(f"  スライド {slide_idx} を処理中...")
        
        # スライド内のすべての図形を処理
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                # テキストボックスやテキスト図形
                if hasattr(shape, "text_frame"):
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            original_text = run.text
                            translated_text = translate_text(original_text, translation_dict)
                            if original_text != translated_text:
                                run.text = translated_text
                                translated_count += 1
                                print(f"    ✓ 翻訳: '{original_text}' → '{translated_text}'")
            
            # テーブルの場合
            if shape.has_table:
                table = shape.table
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.text_frame.paragraphs:
                            for run in paragraph.runs:
                                original_text = run.text
                                translated_text = translate_text(original_text, translation_dict)
                                if original_text != translated_text:
                                    run.text = translated_text
                                    translated_count += 1
                                    print(f"    ✓ テーブル翻訳: '{original_text}' → '{translated_text}'")
    
    # ファイルを保存
    print(f"\n✓ 翻訳完了: {translated_count} 箇所")
    print(f"ファイルを保存中: {output_path}")
    prs.save(output_path)
    print(f"✓ 保存完了")

def translate_text(text, dictionary):
    """テキストを翻訳（辞書を使用）"""
    if not text:
        return text
    
    result = text
    
    # 長い文字列から短い文字列の順で置換（より正確な翻訳のため）
    sorted_keys = sorted(dictionary.keys(), key=len, reverse=True)
    
    for jp_text in sorted_keys:
        if jp_text in result:
            en_text = dictionary[jp_text]
            result = result.replace(jp_text, en_text)
    
    return result

if __name__ == "__main__":
    input_file = "SP-3000-148_Inverter_BlockDiagram2.pptx"
    output_file = "SP-3000-148_Inverter_BlockDiagram2_en.pptx"
    
    print("="*60)
    print("SP-3000-148 PowerPoint 日本語→英語変換")
    print("="*60)
    print()
    
    try:
        translate_presentation(input_file, output_file)
        print()
        print("="*60)
        print(f"✓ 英語版の作成が完了しました")
        print(f"  入力: {input_file}")
        print(f"  出力: {output_file}")
        print("="*60)
    except FileNotFoundError:
        print(f"✗ エラー: ファイルが見つかりません: {input_file}")
        print("  スクリプトと同じディレクトリに配置してください")
    except Exception as e:
        print(f"✗ エラーが発生しました: {e}")
