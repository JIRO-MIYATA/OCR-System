import streamlit as st
import pandas as pd
import docx
import io
import os
import json
from datetime import datetime
from PIL import Image
from dotenv import load_dotenv
from google import genai
from google.genai import types

# .envファイルから環境変数をロード
load_dotenv()

# ---------------------------------------------------------
# ページ初期設定
# ---------------------------------------------------------
st.set_page_config(
    page_title="AI OCR & データ抽出システム",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------
# テーマ設定とカスタムCSSの適用
# ---------------------------------------------------------
if "theme" not in st.session_state:
    st.session_state.theme = "light"

def toggle_theme():
    st.session_state.theme = "dark" if st.session_state.theme == "light" else "light"

IS_DARK = st.session_state.theme == "dark"

# CSS変数の定義
if IS_DARK:
    bg_color = "#09090b"
    bg_subtle = "#0c0c0f"
    card_color = "#0c0c0f"
    border_color = "#1e1e24"
    border_subtle = "#16161a"
    text_color = "#fafafa"
    text_muted = "#a1a1aa"
    shadow = "none"
else:
    bg_color = "#ffffff"
    bg_subtle = "#f9fafb"
    card_color = "#ffffff"
    border_color = "#e4e4e7"
    border_subtle = "#f0f0f2"
    text_color = "#09090b"
    text_muted = "#71717a"
    shadow = "0 1px 3px rgba(0,0,0,0.04), 0 1px 2px rgba(0,0,0,0.03)"

custom_css = f"""
<style>
    /* 全体背景とフォント設定 */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"], .main, .block-container, section[data-testid="stMain"] {{
        background-color: {bg_color} !important;
        color: {text_color} !important;
        font-family: 'DM Sans', -apple-system, sans-serif !important;
    }}
    .block-container {{
        padding: 2rem 2.5rem 3rem !important;
        max-width: 1360px !important;
    }}
    
    /* カード風コンテナ */
    .custom-card {{
        background-color: {card_color};
        border: 1px solid {border_color};
        border-radius: 10px;
        padding: 1.5rem;
        box-shadow: {shadow};
        margin-bottom: 1.25rem;
    }}
    
    /* ヘッダーブランド */
    .brand {{
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 1rem;
    }}
    .brand-title {{
        font-size: 1.75rem;
        font-weight: 700;
        color: {text_color};
        letter-spacing: -0.03em;
    }}
    .brand-subtitle {{
        font-size: 0.9rem;
        color: {text_muted};
        margin-top: -0.5rem;
        margin-bottom: 1.5rem;
    }}

    /* HTMLテーブルのスタイリング */
    .data-table {{
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        font-size: 0.85rem;
        margin-top: 1rem;
    }}
    .data-table th {{
        text-align: left;
        padding: 0.75rem 1rem;
        color: {text_muted};
        font-weight: 600;
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        border-bottom: 2px solid {border_color};
        background-color: {bg_subtle};
    }}
    .data-table td {{
        padding: 0.75rem 1rem;
        color: {text_color};
        border-bottom: 1px solid {border_subtle};
    }}
    .data-table tr:hover td {{
        background-color: {bg_subtle};
    }}
    
    /* バッジ */
    .badge {{
        display: inline-block;
        padding: 2px 8px;
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: 500;
    }}
    .badge-success {{
        color: #16a34a;
        background-color: rgba(22, 163, 74, 0.1);
    }}
    .badge-info {{
        color: #2563eb;
        background-color: rgba(37, 99, 235, 0.1);
    }}

    /* ファイルアップローダー（ドラッグ＆ドロップゾーン）のカスタマイズ */
    [data-testid="stFileUploaderDropzone"] {{
        border: 2px dashed {border_color} !important;
        background-color: {bg_subtle} !important;
        border-radius: 12px !important;
        padding: 2.5rem 2rem !important;
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: center !important;
        text-align: center !important;
        min-height: 180px !important;
        transition: all 0.2s ease-in-out !important;
        cursor: pointer !important;
    }}
    [data-testid="stFileUploaderDropzone"]:hover {{
        border-color: #2563eb !important;
        background-color: rgba(37, 99, 235, 0.04) !important;
    }}
    
    /* アップローダー内のボタン（Upload / Browse files）のカスタマイズ */
    [data-testid="stFileUploaderDropzone"] button {{
        background-color: #2563eb !important;
        color: #ffffff !important;
        border: none !important;
        padding: 0.5rem 1.5rem !important;
        border-radius: 6px !important;
        font-weight: 500 !important;
        margin-top: 0.75rem !important;
        box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.1) !important;
        transition: background-color 0.2s !important;
    }}
    [data-testid="stFileUploaderDropzone"] button:hover {{
        background-color: #1d4ed8 !important;
    }}
    
    /* ドロップゾーン内にアップロードを促すアイコンとテキストを追加 */
    [data-testid="stFileUploaderDropzone"]::before {{
        content: "📥" !important;
        font-size: 2.5rem !important;
        margin-bottom: 0.5rem !important;
    }}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ---------------------------------------------------------
# ファイル名・パス定義
# ---------------------------------------------------------
EXCEL_FILE = "extracted_data.xlsx"

# ---------------------------------------------------------
# 解析ヘルパー関数
# ---------------------------------------------------------



def extract_text_from_word(file_bytes):
    """Wordファイルからテキストを抽出する"""
    try:
        doc = docx.Document(io.BytesIO(file_bytes))
        full_text = []
        # 段落の読み込み
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text)
        # 表の読み込み
        for table in doc.tables:
            for row in table.rows:
                row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if row_text:
                    full_text.append(" | ".join(row_text))
        return "\n".join(full_text)
    except Exception as e:
        st.error(f"Wordファイルの解析中にエラーが発生しました: {e}")
        return ""

def extract_text_from_excel(file_bytes):
    """Excelファイルからテキスト（Markdownの表形式）を抽出する"""
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        full_text = []
        for sheet_name in xl.sheet_names:
            df = xl.parse(sheet_name)
            if not df.empty:
                full_text.append(f"### シート: {sheet_name}")
                # 空白の行・列をドロップし、文字列変換
                df_clean = df.dropna(how='all').fillna("")
                markdown_table = df_clean.to_markdown(index=False)
                full_text.append(markdown_table)
        return "\n\n".join(full_text)
    except Exception as e:
        st.error(f"Excelファイルの解析中にエラーが発生しました: {e}")
        return ""

# ---------------------------------------------------------
# Gemini API 連携ロジック
# ---------------------------------------------------------

def extract_data_with_gemini(api_key, contents, fields):
    """Gemini APIを使用して、テキストまたは画像から構造化されたJSONデータを抽出する"""
    try:
        # 最新のgoogle-genaiクライアントを初期化
        client = genai.Client(api_key=api_key)
        
        # 動的にJSONスキーマを構築
        properties = {
            field: types.Schema(
                type=types.Type.STRING, 
                description=f"ドキュメントから'{field}'に該当する情報を抽出してください。情報がない場合はnullにしてください。"
            ) 
            for field in fields
        }
        
        schema = types.Schema(
            type=types.Type.OBJECT,
            properties=properties,
            required=fields
        )
        
        fields_desc = ", ".join([f'"{f}"' for f in fields])
        prompt = f"""
        提供されたドキュメント（画像またはテキスト）から、以下の項目を正確に抽出してください。
        抽出項目: {fields_desc}
        
        【抽出ルール】
        1. 金額や日付などは、可能な限りクリーンな形式（金額ならカンマなし数値、日付ならYYYY-MM-DD形式など）に整形して抽出してください。
        2. ドキュメント内に明確な情報がない項目は null にしてください。
        """
        
        # リクエストコンテンツの構築
        request_contents = []
        if isinstance(contents, list):
            request_contents.extend(contents)
        else:
            request_contents.append(contents)
        
        request_contents.append(prompt)
        
        # API呼び出し (gemini-2.5-flashを使用)
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=request_contents,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=schema,
                temperature=0.1
            )
        )
        
        # 結果のパース
        return json.loads(response.text)
        
    except Exception as e:
        raise RuntimeError(f"Gemini APIによる解析中にエラーが発生しました: {e}")

# ---------------------------------------------------------
# Excel保存・読み込みロジック
# ---------------------------------------------------------

def load_database():
    """蓄積されたExcelデータをロードする"""
    if os.path.exists(EXCEL_FILE):
        try:
            return pd.read_excel(EXCEL_FILE)
        except Exception as e:
            st.error(f"データベースファイルの読み込みに失敗しました: {e}")
            return pd.DataFrame()
    return pd.DataFrame()

def save_record_to_excel(record):
    """抽出した1つのレコードをExcelに追記・保存する"""
    df_new = pd.DataFrame([record])
    if os.path.exists(EXCEL_FILE):
        try:
            df_existing = pd.read_excel(EXCEL_FILE)
            # 既存の列構成と合わせつつ結合
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        except Exception as e:
            st.warning(f"既存のExcelの読み込みに失敗したため、新規に作成します: {e}")
            df_combined = df_new
    else:
        df_combined = df_new
        
    # 保存
    df_combined.to_excel(EXCEL_FILE, index=False)
    return df_combined

def reset_database():
    """データベースファイルを削除する"""
    if os.path.exists(EXCEL_FILE):
        os.remove(EXCEL_FILE)
        return True
    return False

# ---------------------------------------------------------
# UI 構築
# ---------------------------------------------------------

# ヘッダー
head_left, head_right = st.columns([9, 2])
with head_left:
    st.markdown("""
    <div class="brand">
        <span class="brand-title">📄 Smart Document OCR</span>
    </div>
    <div class="brand-subtitle">
        PDF、Word、Excel、画像をAIで解析し、Excelデータベースへ自動蓄積します。
    </div>
    """, unsafe_allow_html=True)
with head_right:
    theme_label = "☀️ Light" if IS_DARK else "🌙 Dark"
    st.button(theme_label, on_click=toggle_theme, use_container_width=True, key="theme_toggle_btn")

# サイドバー設定
st.sidebar.markdown("### ⚙️ システム設定")

# 1. APIキーの設定 (環境変数優先、なければ入力フォーム)
env_api_key = os.getenv("GEMINI_API_KEY")
if env_api_key:
    api_key = env_api_key
    st.sidebar.success("🔑 Gemini APIキーを環境変数からロードしました")
else:
    api_key = st.sidebar.text_input("🔑 Gemini APIキーを入力してください", type="password")
    if not api_key:
        st.sidebar.warning("⚠️ APIキーが設定されていません。")

# 2. 抽出項目の設定
st.sidebar.markdown("#### 📌 抽出する項目")
st.sidebar.info("抽出したい項目名をカンマ区切りで入力してください。")
default_fields = "日付, 取引先, 金額, 品目"
fields_input = st.sidebar.text_area("抽出項目定義", value=default_fields, height=100)
fields = [f.strip() for f in fields_input.split(",") if f.strip()]

# 3. データベース初期化
st.sidebar.markdown("#### 🗑️ データの初期化")
if st.sidebar.button("データベースをリセット", type="secondary", use_container_width=True):
    if reset_database():
        st.sidebar.success("データベースをリセットしました")
        st.rerun()
    else:
        st.sidebar.info("リセットするデータベースが存在しません")

# メインパネルのレイアウト
upload_tab, db_tab = st.tabs(["📤 ファイルのアップロードと解析", "📊 蓄積データの一覧"])

# ---------------------------------------------------------
# タブ1: アップロードと解析
# ---------------------------------------------------------
with upload_tab:
    st.markdown('<div class="custom-card">', unsafe_allow_html=True)
    st.subheader("ファイルのアップロード")
    uploaded_files = st.file_uploader(
        "PDF、Word、Excel、JPEG、PNGファイルをアップロードしてください（複数一括可能）",
        type=["pdf", "docx", "xlsx", "xls", "jpg", "jpeg", "png"],
        accept_multiple_files=True
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_files:
        st.markdown(f"**アップロードされたファイル数: {len(uploaded_files)} 件**")
        
        if not api_key:
            st.error("❌ Gemini APIキーが設定されていません。サイドバーまたは環境変数で設定してください。")
        else:
            # 抽出ボタン
            if st.button("🚀 AIで一括解析を実行", type="primary"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # 結果を一時保存するリスト
                results = []
                
                for idx, file in enumerate(uploaded_files):
                    status_text.write(f"【{file.name}】を処理中... ({idx+1}/{len(uploaded_files)})")
                    file_ext = os.path.splitext(file.name)[1].lower()
                    file_bytes = file.read()
                    
                    contents = None
                    
                    try:
                        # ファイル形式に応じたパース処理
                        if file_ext in [".jpg", ".jpeg", ".png"]:
                            # 画像としてそのまま
                            contents = [Image.open(io.BytesIO(file_bytes))]
                        elif file_ext == ".pdf":
                            # PDFは直接バイナリデータをPartとして作成し、Geminiに渡す
                            contents = types.Part.from_bytes(
                                data=file_bytes,
                                mime_type="application/pdf"
                            )
                        elif file_ext == ".docx":
                            # Wordのテキスト抽出
                            text = extract_text_from_word(file_bytes)
                            if not text:
                                raise ValueError("Wordファイルからテキストを抽出できませんでした。")
                            contents = text
                        elif file_ext in [".xlsx", ".xls"]:
                            # Excelのテキスト（Markdownテーブル）抽出
                            text = extract_text_from_excel(file_bytes)
                            if not text:
                                raise ValueError("Excelファイルからデータを抽出できませんでした。")
                            contents = text
                        
                        # Gemini API を呼び出してデータを抽出
                        if contents:
                            extracted_data = extract_data_with_gemini(api_key, contents, fields)
                            
                            # メタデータを追加して結果を作成
                            record = {
                                "処理日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "元ファイル名": file.name,
                                "ファイル形式": file_ext.upper().replace(".", ""),
                                **extracted_data
                            }
                            
                            # データベースに保存
                            save_record_to_excel(record)
                            results.append((file.name, "成功", record))
                        else:
                            results.append((file.name, "失敗 (解析可能なコンテンツがありません)", None))
                            
                    except Exception as e:
                        st.error(f"エラー: {file.name} の処理中にエラーが発生しました: {e}")
                        results.append((file.name, f"失敗 ({str(e)})", None))
                    
                    # プログレスバー更新
                    progress_bar.progress((idx + 1) / len(uploaded_files))
                
                status_text.success("✨ すべてのファイルの処理が完了しました！")
                
                # 処理結果のサマリーを表示
                st.markdown("### 📋 処理サマリー")
                summary_data = []
                for fname, status, record in results:
                    summary_data.append({
                        "ファイル名": fname,
                        "ステータス": status,
                        "抽出結果": json.dumps(record, ensure_ascii=False) if record else "-"
                    })
                st.dataframe(pd.DataFrame(summary_data), use_container_width=True)

# ---------------------------------------------------------
# タブ2: 蓄積データの一覧
# ---------------------------------------------------------
with db_tab:
    df_db = load_database()
    
    if df_db.empty:
        st.info("📂 まだ蓄積されたデータはありません。ファイルを解析するとここに追加されます。")
    else:
        st.markdown('<div class="custom-card">', unsafe_allow_html=True)
        st.subheader("蓄積データ")
        
        # カラム順序を整理 (処理日時、ファイル名、ファイル形式を左側に配置)
        meta_cols = ["処理日時", "元ファイル名", "ファイル形式"]
        other_cols = [col for col in df_db.columns if col not in meta_cols]
        final_cols = [col for col in meta_cols if col in df_db.columns] + other_cols
        df_display = df_db[final_cols]
        
        # Streamlit標準のインタラクティブテーブルで表示
        st.dataframe(df_display, use_container_width=True)
        
        # Excelファイルのダウンロードボタン
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                label="📥 Excelファイルとしてダウンロード",
                data=f,
                file_name="ocr_extracted_database.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        st.markdown('</div>', unsafe_allow_html=True)
