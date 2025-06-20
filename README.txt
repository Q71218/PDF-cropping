PDF 多選區裁切工具（10x15cm）  設計：DY
----------------------------

【工具說明】
本工具可從 PDF 檔案中裁切多個自定區域，並匯出成固定尺寸 (10x15 公分) 的新 PDF 檔案。
支援多頁批次裁切，適用於照片列印、表單剪裁、證件分割等需求。
框選的順序 為 新 PDF 排列順序

----------------------------
【使用方式】

1. 開啟 PDF
   - 點擊「開啟 PDF」按鈕，選擇欲處理的 PDF 檔案。
   - 若檔案有密碼保護，預設會先嘗試密碼 `66608251`，若失敗會提示輸入密碼。
★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
★★★★★★ 可自行於程式碼修改預設密碼 ★★★★★★
★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

2. 選擇裁切區域
   - 使用滑鼠左鍵在畫面上拖曳，即可選擇欲裁切的區域。
   - 可多次選取，選區會依序編號顯示。
   - 若需重新選取，請點擊「清除選取區」。

3. 多頁裁切（可選）
   - 勾選「批次多頁處理」：會對整本 PDF 中的每一頁使用相同選區進行裁切。
   - 不勾選：僅針對當前頁面裁切。

4. 切換頁面
   - 使用「上一頁」與「下一頁」瀏覽其他頁面。
   - 切換頁面後會清除當前的選取區（避免誤用）。

5. 匯出裁切結果
   - 點擊「匯出裁切區」開始匯出。
   - 每一個選區會匯出為一頁，結果為 10x15 cm 的 PDF 文件。
   - 新檔會儲存在原始 PDF 檔案相同資料夾，檔名類似：
     `原檔名_已裁切_3頁.pdf`

----------------------------
【視窗說明】
- 工具列上方：基本操作按鈕
- 中央灰色畫布：顯示 PDF 頁面與選取區域

----------------------------
【系統需求】
- Windows / macOS / Linux（需支援 Python 3）
- 依賴套件：
  - PyMuPDF (fitz)
  - Pillow
  - tkinter（Python 內建）

----------------------------
【安裝套件指令】
請在 命令提示字元 輸入：
pip install pymupdf pillow

【封裝成EXE的指令】
pyinstaller --onefile --windowed --name pdf-cropper pdf-cropper.py

----------------------------
【注意事項】
- 匯出時使用解析度 300dpi，品質足以列印。
- 本工具不支援旋轉頁面、密碼加密或 OCR 文字處理。
- 裁切區座標以原始 PDF 為基準，請避免選取超出頁面範圍。

----------------------------
【熱感印表機列印事項】
以Xprinter XP-470E 熱感印表機為列
標簽樣式選擇內建
4*6(101.6mm*152.4mm)

----------------------------
Design by DY
感謝使用本工具！
