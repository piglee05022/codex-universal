# 法律文書產製模組（A01）

此模組可自動產製民事、刑事及行政訴訟等法院文書，支援繁體中文格式、標楷體字型、四邊 2.5 公分邊界與 1.5 倍行距。當安裝 `docx2pdf` 時，可同時輸出 Word 與 PDF 檔案。

## 主要功能

- 自動依據案件資料產製起訴狀、答辯狀等常見文書。
- 整合證據清單與附件，段落標號自動化。
- 依案件所屬法院與層級套用不同封面格式（可擴充）。

## 安裝相依套件

```bash
pip install python-docx docx2pdf
```

## 使用方式

可直接執行 `example.py` 以產生範例文書：

```bash
python3 -m legal_module.example
```

或在程式中呼叫 `create_filing`：

```python
from legal_module.filing import create_filing

case_info = {
    'case_number': '臺北地方法院114年度訴字第1234號',
    'parties': '原告：李灃祐  被告：新鑫公司',
    'court': '臺灣臺北地方法院',
    'claims': '請求確認本票債權不存在，並請求返還不當得利新台幣1,000,000元。',
    'facts': '原告與被告簽署車輛分期契約，惟車輛自始未交付，卻遭告提出票據裁定……',
    'laws': ['民法第184條', '票據法第17條', '最高法院111年度台上字第3208號判決'],
    'evidence': [
        {'id': '乙1', 'summary': 'LINE對話紀錄，顯示告知車輛尚未交付'},
        {'id': '乙2', 'summary': '川立公司匯款憑證，顯示資金流向'}
    ],
    'attachments': [
        {'id': '附件一', 'description': '車輛分期契約影本'}
    ]
}

create_filing(case_info, '法律文書_起訴狀.docx', pdf_path='法律文書_起訴狀.pdf')
```

執行後將產生 `法律文書_起訴狀.docx` 及 `法律文書_起訴狀.pdf` 檔案，可用於列印或上傳 e訴訟平台。

