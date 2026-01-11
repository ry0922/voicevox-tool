import io
import os
import wave
import struct
from dotenv import load_dotenv
import requests
import gspread
from google.oauth2.service_account import Credentials

# =========================
# 設定項目
# =========================
load_dotenv()

# VOICEVOX エンジンのURL
VOICEVOX_URL = "http://127.0.0.1:50021"
SPEAKER_ID = 8  # 好きな話者IDに変更

# 無音の長さ（秒）
SILENCE_SECONDS = 30

# Google Sheets 設定
SERVICE_ACCOUNT_JSON = os.getenv("SERVICE_ACCOUNT_JSON")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
RANGE_NAME = "シート1!A2:A"  # A列の2行目以降を読む例

# 出力ファイル名
OUTPUT_WAV = "output_with_silence.wav"


# =========================
# Google Sheets からテキスト取得
# =========================

def load_texts_from_spreadsheet():
    scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
    ]
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_JSON,
        scopes=scopes,
    )
    gc = gspread.authorize(credentials)
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.worksheet(RANGE_NAME.split("!")[0])  # シート名部分だけ取り出し

    # A列の2行目以降の値を取得（空行は除外）
    # RANGE_NAME の列/行指定を使わず、シンプルに A列全部読む例にしておく
    values = worksheet.col_values(1)  # A列
    # 1行目はヘッダ想定ならスキップ
    if values:
        values = values[1:]
    texts = [v.strip() for v in values if v.strip()]
    return texts



# =========================
# VOICEVOX で音声生成
# =========================

def synthesize_voicevox(text, speaker=SPEAKER_ID):
    """
    VOICEVOX に text を投げて wav のバイナリを返す
    """
    # audio_query を取得
    query_params = {"text": text, "speaker": speaker}
    r = requests.post(f"{VOICEVOX_URL}/audio_query", params=query_params)
    r.raise_for_status()
    audio_query = r.json()

    # synthesis で音声生成
    r2 = requests.post(
        f"{VOICEVOX_URL}/synthesis",
        params={"speaker": speaker},
        json=audio_query,
    )
    r2.raise_for_status()
    wav_bytes = r2.content
    return wav_bytes


# =========================
# 無音のwavフレーム生成
# =========================

def create_silence_frames(n_channels, sampwidth, framerate, seconds):
    """
    指定したフォーマットで seconds 秒分の無音の生フレーム(bytes)を作る
    """
    n_frames = int(framerate * seconds)
    # PCM16bit前提（sampwidth=2）の場合
    if sampwidth != 2:
        raise ValueError("このサンプルは16bit PCM (sampwidth=2) のみ対応です。")

    # 全て0のサンプル（無音）
    silence = struct.pack("<h", 0) * n_channels * n_frames
    return silence


# =========================
# 複数wav + 無音を結合
# =========================

def concat_wavs_with_silence(wav_list, silence_seconds, output_wav):
    """
    wav_list: 各要素が wav バイナリ(bytes) のリスト
    silence_seconds: 各発言の間に入れる無音の秒数
    output_wav: 出力ファイルパス
    """
    # まず最初のwavのフォーマットを基準にする
    first_wav = wav_list[0]
    with wave.open(io.BytesIO(first_wav), "rb") as w:
        n_channels = w.getnchannels()
        sampwidth = w.getsampwidth()
        framerate = w.getframerate()
        comptype = w.getcomptype()
        compname = w.getcompname()

    # 出力ファイルを作成
    with wave.open(output_wav, "wb") as out_w:
        out_w.setnchannels(n_channels)
        out_w.setsampwidth(sampwidth)
        out_w.setframerate(framerate)
        out_w.setcomptype(comptype, compname)

        silence_frames = create_silence_frames(
            n_channels, sampwidth, framerate, silence_seconds
        )

        for i, wav_bytes in enumerate(wav_list):
            # 音声本体を書き込み
            with wave.open(io.BytesIO(wav_bytes), "rb") as w:
                # 念のためフォーマットが同じかチェック
                if (
                    w.getnchannels() != n_channels
                    or w.getsampwidth() != sampwidth
                    or w.getframerate() != framerate
                ):
                    raise ValueError("wav のフォーマットが一致していません。")
                frames = w.readframes(w.getnframes())
                out_w.writeframes(frames)

            # 最後の発言以外なら無音を挟む
            if i != len(wav_list) - 1:
                out_w.writeframes(silence_frames)


# =========================
# メイン処理
# =========================

def main():
    # 1. シートからテキスト一覧取得
    texts = load_texts_from_spreadsheet()
    # texts = [
    #     "こんにちは、これはテストです。",
    #     "VOICEVOXを使って音声を生成しています。",
    #     "各発言の間に無音を挿入します。"
    # ]  # テスト用ダミーデータ
    if not texts:
        print("スプレッドシートにテキストがありません。")
        return

    print(f"{len(texts)} 件のテキストを読み込みました。音声を生成します…")

    # 2. 各テキストを VOICEVOX で音声生成
    wav_list = []
    for i, text in enumerate(texts, start=1):
        print(f"[{i}/{len(texts)}] 合成中: {text}")
        wav_bytes = synthesize_voicevox(text)
        wav_list.append(wav_bytes)

    # 3. 無音を挟んで1つのwavに結合
    print("結合中（各発言の間に30秒の無音を挿入）…")
    concat_wavs_with_silence(wav_list, SILENCE_SECONDS, OUTPUT_WAV)

    print(f"完了しました！ 出力ファイル: {OUTPUT_WAV}")


if __name__ == "__main__":
    main()
