#!/usr/bin/env python
# coding: utf-8

# In[4]:


# ============================================
# 程式名稱：合併 Shopify 和 Pinkoi 訂單表為月度財務報表
# 檔案路徑：
#   - Shopify: C:\Users\MI\Desktop\2026-月度財務報表\01\Shopify-Orders_計算版-V3.xlsx
#   - Pinkoi:  C:\Users\MI\Desktop\2026-月度財務報表\01\202601-Pinkoi_orders.xlsx
# 輸出：月度財務報表_202601.xlsx
# 工作表：
#   - 月度統計
#   - 渠道對比
#   - Shopify訂單明細
#   - Pinkoi訂單明細
# ============================================

import pandas as pd
import os
import numpy as np

print("=" * 60)
print("📦 開始合併 Shopify 和 Pinkoi 訂單表...")
print("=" * 60)

# === 1. 設定檔案路徑 ===
folder_path = r'C:\Users\MI\Desktop\2026-月度財務報表\01'
shopify_path = os.path.join(folder_path, 'Shopify-Orders_計算版-V3.xlsx')
pinkoi_path = os.path.join(folder_path, '202601-Pinkoi_orders.xlsx')
output_path = os.path.join(folder_path, '月度財務報表_202601.xlsx')

# === 2. 檢查檔案是否存在 ===
if not os.path.exists(shopify_path):
    print(f"❌ 錯誤：找不到 Shopify 訂單表")
    exit()

if not os.path.exists(pinkoi_path):
    print(f"❌ 錯誤：找不到 Pinkoi 訂單表")
    exit()

# === 3. 讀取檔案 ===
print(f"\n📂 正在讀取 Shopify 訂單表...")
shopify = pd.read_excel(shopify_path)
print(f"✅ Shopify：{len(shopify)} 行，{len(shopify.columns)} 欄")

print(f"\n📂 正在讀取 Pinkoi 訂單表...")
pinkoi = pd.read_excel(pinkoi_path)
print(f"✅ Pinkoi：{len(pinkoi)} 行，{len(pinkoi.columns)} 欄")

# === 4. 為兩個表加上渠道標記 ===
shopify['渠道'] = 'Shopify'
pinkoi['渠道'] = 'Pinkoi'

# === 5. 標準化 Shopify 欄位 ===
print("\n🔄 標準化 Shopify 欄位...")

shopify_std = pd.DataFrame({
    '渠道': shopify['渠道'],
    '訂單編號': shopify['Order No'],
    '訂單日期': shopify['Created at'],
    '客戶名稱': shopify['Customer Name'],
    '商品名稱': shopify['Product Name'],
    '數量': pd.to_numeric(shopify['Quantity'], errors='coerce').fillna(0),
    '單價': pd.to_numeric(shopify['Selling Price'], errors='coerce').fillna(0),
    '總金額': pd.to_numeric(shopify['Total'], errors='coerce').fillna(0),
    '折扣': pd.to_numeric(shopify['Discount Amount'], errors='coerce').fillna(0),
    '分攤後金額': pd.to_numeric(shopify['分攤後金額'], errors='coerce').fillna(0),
    '分攤後折扣': pd.to_numeric(shopify['分攤後折扣'], errors='coerce').fillna(0),
    '成本': pd.to_numeric(shopify['Cost  (unit)'], errors='coerce').fillna(0),
    '單件利潤': pd.to_numeric(shopify['Profit (unit)'], errors='coerce').fillna(0),
    '總成本': pd.to_numeric(shopify['Total Cost'], errors='coerce').fillna(0),
    '總利潤': pd.to_numeric(shopify['Total Profit'], errors='coerce').fillna(0),
    '利潤率': pd.to_numeric(shopify[' Gross Profit Margin'], errors='coerce').fillna(0)
})

# === 6. 標準化 Pinkoi 欄位 ===
print("🔄 標準化 Pinkoi 欄位...")

pinkoi_std = pd.DataFrame({
    '渠道': pinkoi['渠道'],
    '訂單編號': pinkoi['訂單編號'],
    '訂單日期': pd.to_datetime(pinkoi['訂單成立日期'], errors='coerce'),
    '客戶名稱': pinkoi['買家'],
    '商品名稱': pinkoi['購買品項'],
    '數量': pd.to_numeric(pinkoi['數量'], errors='coerce').fillna(0),
    '單價': pd.to_numeric(pinkoi['商品單價'], errors='coerce').fillna(0),
    '總金額': pd.to_numeric(pinkoi['總金額'], errors='coerce').fillna(0),
    '折扣': pd.to_numeric(pinkoi['折抵'], errors='coerce').fillna(0),
    '分攤後金額': 0,  # Pinkoi 沒有分攤後金額，先用0
    '分攤後折扣': 0,  # Pinkoi 沒有分攤後折扣，先用0
    '成本': 0,  # Pinkoi 沒有成本資料
    '單件利潤': 0,  # Pinkoi 沒有利潤資料
    '總成本': 0,
    '總利潤': 0,
    '利潤率': 0
})

# === 7. 計算 Pinkoi 的衍生欄位 ===
print("🔄 計算 Pinkoi 衍生欄位...")

# Pinkoi 的商品原始金額
pinkoi_std['商品原始金額'] = pinkoi_std['數量'] * pinkoi_std['單價']

# Pinkoi 的實際金額（如果沒有分攤後金額，就用商品原始金額）
pinkoi_std['實際金額'] = pinkoi_std['商品原始金額']

# Pinkoi 的總利潤（沒有成本，所以利潤 = 實際金額）
pinkoi_std['總利潤'] = pinkoi_std['實際金額']
pinkoi_std['利潤率'] = 100.0  # 沒有成本，利潤率100%

# 驗證 Pinkoi 的總金額是否等於商品原始金額（考慮折扣）
for idx, row in pinkoi_std.iterrows():
    if abs(row['商品原始金額'] - row['折扣'] - row['總金額']) > 1:
        print(f"   ⚠️ 訂單 {row['訂單編號']} 金額不一致：商品原價 {row['商品原始金額']} - 折扣 {row['折扣']} ≠ 總金額 {row['總金額']}")

# === 8. 計算 Shopify 的衍生欄位 ===
print("🔄 計算 Shopify 衍生欄位...")

# Shopify 的商品原始金額
shopify_std['商品原始金額'] = shopify_std['數量'] * shopify_std['單價']

# Shopify 的實際金額（優先使用分攤後金額）
shopify_std['實際金額'] = shopify_std.apply(
    lambda row: row['分攤後金額'] if row['分攤後金額'] > 0 else row['商品原始金額'],
    axis=1
)

# === 9. 定義統一的欄位順序 ===
final_cols = [
    '渠道', '訂單編號', '訂單日期', '客戶名稱', '商品名稱',
    '數量', '單價', '商品原始金額', '折扣', '分攤後金額', '分攤後折扣',
    '實際金額', '總金額', '成本', '總成本', '單件利潤', '總利潤', '利潤率'
]

# === 10. 分別處理兩個渠道 ===
print("\n📊 處理 Shopify 訂單明細...")
shopify_final = shopify_std[final_cols].copy()
shopify_final = shopify_final.sort_values(['訂單日期', '訂單編號'])

print("📊 處理 Pinkoi 訂單明細...")
pinkoi_final = pinkoi_std[final_cols].copy()
pinkoi_final = pinkoi_final.sort_values(['訂單日期', '訂單編號'])

# === 11. 合併用於統計（不輸出）===
combined = pd.concat([shopify_final, pinkoi_final], ignore_index=True, sort=False)

# === 12. 生成月度統計 ===
print("\n📊 生成月度統計...")

# 整體統計
total_orders = combined['訂單編號'].nunique()
total_items = len(combined)
total_sales = combined[combined['總金額'] > 0]['總金額'].sum()
total_actual = combined['實際金額'].sum()
total_discount = combined['折扣'].sum()
total_profit = combined['總利潤'].sum()

# 渠道統計
channel_stats = combined.groupby('渠道').agg({
    '訂單編號': 'nunique',
    '實際金額': 'sum',
    '折扣': 'sum',
    '總利潤': 'sum'
}).round(2)
channel_stats.columns = ['訂單數', '營業額', '折扣總額', '總利潤']
channel_stats['佔比'] = (channel_stats['營業額'] / total_actual * 100).round(1).astype(str) + '%'
channel_stats['利潤率'] = (channel_stats['總利潤'] / channel_stats['營業額'] * 100).round(1).astype(str) + '%'

# 建立統計表
stats_data = {
    '統計項目': [
        '📊 整體概覽',
        '總訂單數',
        '總商品明細數',
        '總營業額',
        '總折扣金額',
        '總利潤',
        '平均利潤率',
        '平均客單價',
        '',
        '📈 渠道分析',
        'Shopify 訂單數',
        'Shopify 營業額',
        'Shopify 佔比',
        'Shopify 利潤',
        'Shopify 利潤率',
        'Pinkoi 訂單數',
        'Pinkoi 營業額',
        'Pinkoi 佔比',
        'Pinkoi 利潤',
        'Pinkoi 利潤率'
    ],
    '數值': [
        '',
        f"{total_orders} 筆",
        f"{total_items} 筆",
        f"${total_actual:,.2f}",
        f"${total_discount:,.2f}",
        f"${total_profit:,.2f}",
        f"{(total_profit/total_actual*100):.1f}%" if total_actual > 0 else '0%',
        f"${total_actual/total_orders:,.2f}" if total_orders > 0 else '$0',
        '',
        '',
        f"{channel_stats.loc['Shopify', '訂單數'] if 'Shopify' in channel_stats.index else 0} 筆",
        f"${channel_stats.loc['Shopify', '營業額'] if 'Shopify' in channel_stats.index else 0:,.2f}",
        f"{channel_stats.loc['Shopify', '佔比'] if 'Shopify' in channel_stats.index else '0%'}",
        f"${channel_stats.loc['Shopify', '總利潤'] if 'Shopify' in channel_stats.index else 0:,.2f}",
        f"{channel_stats.loc['Shopify', '利潤率'] if 'Shopify' in channel_stats.index else '0%'}",
        f"{channel_stats.loc['Pinkoi', '訂單數'] if 'Pinkoi' in channel_stats.index else 0} 筆",
        f"${channel_stats.loc['Pinkoi', '營業額'] if 'Pinkoi' in channel_stats.index else 0:,.2f}",
        f"{channel_stats.loc['Pinkoi', '佔比'] if 'Pinkoi' in channel_stats.index else '0%'}",
        f"${channel_stats.loc['Pinkoi', '總利潤'] if 'Pinkoi' in channel_stats.index else 0:,.2f}",
        f"{channel_stats.loc['Pinkoi', '利潤率'] if 'Pinkoi' in channel_stats.index else '0%'}"
    ]
}

stats_df = pd.DataFrame(stats_data)

# === 13. 儲存檔案 ===
print(f"\n💾 正在儲存檔案：{output_path}")

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    # 月度統計
    stats_df.to_excel(writer, sheet_name='月度統計', index=False)

    # 渠道對比
    channel_stats.to_excel(writer, sheet_name='渠道對比')

    # Shopify 訂單明細
    shopify_final.to_excel(writer, sheet_name='Shopify訂單明細', index=False)

    # Pinkoi 訂單明細
    pinkoi_final.to_excel(writer, sheet_name='Pinkoi訂單明細', index=False)

print(f"✅ 完成！已儲存為：{output_path}")

# === 14. 顯示摘要 ===
print("\n" + "=" * 60)
print("📊 2026年1月財務摘要")
print("=" * 60)
print(f"\n總訂單數：{total_orders} 筆")
print(f"總營業額：${total_actual:,.2f}")
print(f"總折扣：${total_discount:,.2f}")
print(f"總利潤：${total_profit:,.2f}")
print(f"平均利潤率：{(total_profit/total_actual*100):.1f}%" if total_actual > 0 else "0%")
print(f"平均客單價：${total_actual/total_orders:,.2f}" if total_orders > 0 else "")

print("\n渠道分佈：")
for channel, row in channel_stats.iterrows():
    print(f"\n  {channel}：")
    print(f"    訂單數：{row['訂單數']} 單")
    print(f"    營業額：${row['營業額']:,.2f} ({row['佔比']})")
    print(f"    利潤：${row['總利潤']:,.2f} ({row['利潤率']})")

print(f"\n📋 工作表說明：")
print(f"   1. 月度統計 - 整體財務指標")
print(f"   2. 渠道對比 - Shopify vs Pinkoi 比較")
print(f"   3. Shopify訂單明細 - {len(shopify_final)} 筆明細")
print(f"   4. Pinkoi訂單明細 - {len(pinkoi_final)} 筆明細")

print("\n" + "=" * 60)
print("🎉 完成！")
print("=" * 60)


# In[ ]:




