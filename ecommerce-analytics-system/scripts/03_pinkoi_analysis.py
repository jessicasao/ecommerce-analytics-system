#!/usr/bin/env python
# coding: utf-8

# In[2]:


# ============================================
# 程式名稱：Pinkoi 2025年訂單統計（修正買家計算）
# 檔案路徑：C:\Users\MI\Desktop\Pinkoi_Orders\2025\Pinkoi_2025統計.xlsx
# 修正：買家數量 = 總訂單數（人次），不是不重複人數
# ============================================

import pandas as pd
import numpy as np
from datetime import datetime
import os

print("=" * 60)
print("📊 開始更新 Pinkoi 2025年統計...")
print("=" * 60)

# === 1. 設定檔案路徑 ===
file_path = r'C:\Users\MI\Desktop\Pinkoi_Orders\2025\Pinkoi_2025統計.xlsx'
output_path = file_path  # 直接覆蓋原檔案

# === 2. 檢查檔案是否存在 ===
if not os.path.exists(file_path):
    print(f"❌ 錯誤：找不到檔案")
    print(f"   路徑：{file_path}")
    exit()

# === 3. 讀取 Pinkoi 訂單表 ===
print(f"\n📂 正在讀取：{file_path}")
# 讀取所有工作表
xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names
print(f"✅ 找到工作表：{sheet_names}")

# 讀取 2025訂單明細 工作表
if '2025訂單明細' in sheet_names:
    df = pd.read_excel(file_path, sheet_name='2025訂單明細')
    print(f"✅ 讀取『2025訂單明細』：{len(df)} 行，{len(df.columns)} 欄")
else:
    print(f"❌ 錯誤：找不到『2025訂單明細』工作表")
    exit()

# === 4. 顯示所有欄位，幫助識別 ===
print("\n📋 訂單明細欄位：")
for i, col in enumerate(df.columns):
    print(f"  {i+1:2d}. '{col}'")

# === 5. 找出需要的欄位 ===
print("\n🔍 識別數值欄位...")

# 買家名字欄位
buyer_col = None
possible_buyer_cols = ['買家名字', '買家姓名', '客戶名稱', '客戶姓名', '姓名', 'Billing Name', '收件人']
for col in possible_buyer_cols:
    if col in df.columns:
        buyer_col = col
        break

# 總金額欄位
total_col = None
possible_total_cols = ['總金額', '訂單總額', '總計', 'Total', '訂單金額']
for col in possible_total_cols:
    if col in df.columns:
        total_col = col
        break

# 小計欄位
subtotal_col = None
possible_subtotal_cols = ['小計', '商品金額', 'Subtotal', '商品總額']
for col in possible_subtotal_cols:
    if col in df.columns:
        subtotal_col = col
        break

# 折抵欄位
discount_col = None
possible_discount_cols = ['折抵', '折扣', '優惠', 'Discount', '折抵金額']
for col in possible_discount_cols:
    if col in df.columns:
        discount_col = col
        break

# 運費欄位
shipping_col = None
possible_shipping_cols = ['運費', 'Shipping', '運費金額']
for col in possible_shipping_cols:
    if col in df.columns:
        shipping_col = col
        break

print(f"\n📊 找到的欄位：")
print(f"   - 買家名字：{buyer_col if buyer_col else '❌ 未找到'}")
print(f"   - 總金額：{total_col if total_col else '❌ 未找到'}")
print(f"   - 小計：{subtotal_col if subtotal_col else '❌ 未找到'}")
print(f"   - 折抵：{discount_col if discount_col else '❌ 未找到'}")
print(f"   - 運費：{shipping_col if shipping_col else '❌ 未找到'}")

# === 6. 確保數值欄位是數字 ===
print("\n🔄 轉換數值欄位...")

for col in [total_col, subtotal_col, discount_col, shipping_col]:
    if col and col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# === 7. 計算統計數據 ===
print("\n💰 計算統計數據...")

# 總訂單數
total_orders = len(df)

# 買家數量（人次）= 總訂單數
# 因為每一筆訂單就是一個買家（即使同一個人下多單，也算多次）
buyer_count = total_orders

# 不重複買家人數（僅供參考）
if buyer_col:
    unique_buyers = df[buyer_col].dropna().nunique()
else:
    unique_buyers = total_orders

# 總金額
total_amount = df[total_col].sum() if total_col else 0

# 總小計
total_subtotal = df[subtotal_col].sum() if subtotal_col else 0

# 總折抵
total_discount = df[discount_col].sum() if discount_col else 0

# 總運費
total_shipping = df[shipping_col].sum() if shipping_col else 0

# 計算佔比
subtotal_percentage = (total_subtotal / total_amount * 100) if total_amount > 0 else 0
discount_percentage = (total_discount / total_amount * 100) if total_amount > 0 else 0
shipping_percentage = (total_shipping / total_amount * 100) if total_amount > 0 else 0

# 平均客單價
avg_order_value = total_amount / total_orders if total_orders > 0 else 0

# 平均每不重複買家貢獻
avg_per_unique_buyer = total_amount / unique_buyers if unique_buyers > 0 else 0

# 最高/最低單筆金額
max_amount = df[total_col].max() if total_col else 0
min_amount = df[total_col].min() if total_col else 0

# 有折抵的訂單
if discount_col:
    discount_orders = len(df[df[discount_col] > 0])
    discount_order_percentage = (discount_orders / total_orders * 100) if total_orders > 0 else 0
else:
    discount_orders = 0
    discount_order_percentage = 0

# 有運費的訂單
if shipping_col:
    shipping_orders = len(df[df[shipping_col] > 0])
    shipping_order_percentage = (shipping_orders / total_orders * 100) if total_orders > 0 else 0
else:
    shipping_orders = 0
    shipping_order_percentage = 0

# 重複購買分析（如果有人下多單）
if buyer_col and unique_buyers > 0:
    repeat_rate = (total_orders - unique_buyers) / total_orders * 100 if total_orders > 0 else 0
    avg_orders_per_buyer = total_orders / unique_buyers if unique_buyers > 0 else 0
else:
    repeat_rate = 0
    avg_orders_per_buyer = 1

print(f"\n📊 計算結果：")
print(f"   - 總訂單數：{total_orders}")
print(f"   - 買家數量（人次）：{buyer_count}")  # 等於總訂單數
print(f"   - 不重複買家人數：{unique_buyers}")
print(f"   - 平均每人下單次數：{avg_orders_per_buyer:.2f}")
print(f"   - 重複購買率：{repeat_rate:.2f}%")
print(f"   - 總金額：{total_amount:,.2f}")

# === 8. 建立統計表 ===
print("\n📋 建立統計報表...")

stats_data = {
    '統計項目': [
        '📦 訂單概況',
        '總訂單數 (筆)',
        '買家數量 (人次)',
        '不重複買家人數',
        '平均每人下單次數',
        '重複購買率',
        '平均客單價',
        '平均每不重複買家貢獻',
        '',
        '💰 金額分析',
        '總金額',
        '總小計 (商品金額)',
        '總折抵 (折扣/優惠)',
        '總運費',
        '',
        '📊 佔比分析',
        '小計佔總金額比例',
        '折抵佔總金額比例',
        '運費佔總金額比例',
        '',
        '📈 極值分析',
        '最高單筆金額',
        '最低單筆金額',
        '',
        '🏷️ 折抵分析',
        '有折抵的訂單數',
        '折抵訂單佔比',
        '',
        '🚚 運費分析',
        '有運費的訂單數',
        '運費訂單佔比'
    ],
    '數值': [
        '',
        f"{total_orders:,} 筆",
        f"{buyer_count:,} 人次",
        f"{unique_buyers:,} 人",
        f"{avg_orders_per_buyer:.2f} 次",
        f"{repeat_rate:.2f}%",
        f"${avg_order_value:,.2f}",
        f"${avg_per_unique_buyer:,.2f}",
        '',
        '',
        f"${total_amount:,.2f}",
        f"${total_subtotal:,.2f}",
        f"${total_discount:,.2f}",
        f"${total_shipping:,.2f}",
        '',
        '',
        f"{subtotal_percentage:.2f}%",
        f"{discount_percentage:.2f}%",
        f"{shipping_percentage:.2f}%",
        '',
        '',
        f"${max_amount:,.2f}",
        f"${min_amount:,.2f}",
        '',
        '',
        f"{discount_orders:,} 筆",
        f"{discount_order_percentage:.2f}%",
        '',
        '',
        f"{shipping_orders:,} 筆",
        f"{shipping_order_percentage:.2f}%"
    ]
}

stats_df = pd.DataFrame(stats_data)

# === 9. 儲存報表 ===
print(f"\n💾 正在更新統計表：{output_path}")

# 讀取所有現有工作表
with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    # 只替換 '2025統計' 工作表
    stats_df.to_excel(writer, sheet_name='2025統計', index=False)

print(f"✅ 完成！已更新：{output_path}")

# === 10. 顯示摘要 ===
print("\n" + "=" * 60)
print("📊 Pinkoi 2025年統計摘要")
print("=" * 60)

print(f"\n📦 訂單概況：")
print(f"   ├─ 總訂單數：{total_orders:,} 筆")
print(f"   ├─ 買家數量：{buyer_count:,} 人次")  # 修正：這是人次
print(f"   ├─ 不重複買家：{unique_buyers:,} 人")
print(f"   ├─ 平均每人下單：{avg_orders_per_buyer:.2f} 次")
print(f"   ├─ 重複購買率：{repeat_rate:.2f}%")
print(f"   ├─ 平均客單價：${avg_order_value:,.2f}")
print(f"   └─ 平均每不重複買家貢獻：${avg_per_unique_buyer:,.2f}")

print(f"\n💰 金額分析：")
print(f"   ├─ 總金額：${total_amount:,.2f}")
print(f"   ├─ 總小計：${total_subtotal:,.2f}")
print(f"   ├─ 總折抵：-${total_discount:,.2f}")
print(f"   └─ 總運費：+${total_shipping:,.2f}")

print(f"\n📊 佔比分析：")
print(f"   ├─ 小計佔比：{subtotal_percentage:.2f}%")
print(f"   ├─ 折抵佔比：{discount_percentage:.2f}%")
print(f"   └─ 運費佔比：{shipping_percentage:.2f}%")

print(f"\n🏷️ 折抵分析：")
print(f"   ├─ 有折抵訂單：{discount_orders:,} 筆")
print(f"   └─ 折抵訂單佔比：{discount_order_percentage:.2f}%")

print("\n" + "=" * 60)
print("🎉 統計更新完成！")
print("=" * 60)


# In[ ]:





# In[1]:


# ============================================
# 程式名稱：Pinkoi 2025年訂單統計（修正買家欄位）
# 檔案路徑：C:\Users\MI\Desktop\Pinkoi_Orders\2025\Pinkoi_2025統計.xlsx
# 修正：買家欄位是『買家』
# ============================================

import pandas as pd
import numpy as np
from datetime import datetime
import os

print("=" * 60)
print("📊 開始更新 Pinkoi 2025年統計...")
print("=" * 60)

# === 1. 設定檔案路徑 ===
file_path = r'C:\Users\MI\Desktop\Pinkoi_Orders\2025\Pinkoi_2025統計.xlsx'
output_path = file_path

# === 2. 檢查檔案是否存在 ===
if not os.path.exists(file_path):
    print(f"❌ 錯誤：找不到檔案")
    exit()

# === 3. 讀取 Pinkoi 訂單表 ===
print(f"\n📂 正在讀取：{file_path}")
xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names
print(f"✅ 找到工作表：{sheet_names}")

if '2025訂單明細' in sheet_names:
    df = pd.read_excel(file_path, sheet_name='2025訂單明細')
    print(f"✅ 讀取『2025訂單明細』：{len(df)} 行，{len(df.columns)} 欄")
else:
    print(f"❌ 錯誤：找不到『2025訂單明細』工作表")
    exit()

# === 4. 顯示所有欄位 ===
print("\n📋 訂單明細欄位：")
for i, col in enumerate(df.columns):
    print(f"  {i+1:2d}. '{col}'")

# === 5. 找出需要的欄位 ===
print("\n🔍 識別欄位...")

# 買家欄位 - 關鍵修正：加入『買家』
buyer_col = None
possible_buyer_cols = ['買家', '買家名字', '買家姓名', '客戶名稱', '客戶姓名', '姓名', 'Billing Name', '收件人']
for col in possible_buyer_cols:
    if col in df.columns:
        buyer_col = col
        print(f"✅ 找到買家欄位：『{buyer_col}』")
        break

if not buyer_col:
    print("❌ 錯誤：找不到買家欄位！")
    print("請確認以下欄位是否存在：")
    for col in possible_buyer_cols:
        print(f"   - {col}")
    exit()

# 總金額欄位
total_col = '總金額' if '總金額' in df.columns else None

# 小計欄位
subtotal_col = '小計' if '小計' in df.columns else None

# 折抵欄位
discount_col = '折抵' if '折抵' in df.columns else None

# 運費欄位
shipping_col = '運費' if '運費' in df.columns else None

print(f"\n📊 找到的欄位：")
print(f"   - 買家：{buyer_col}")
print(f"   - 總金額：{total_col}")
print(f"   - 小計：{subtotal_col}")
print(f"   - 折抵：{discount_col}")
print(f"   - 運費：{shipping_col}")

# === 6. 確保數值欄位是數字 ===
print("\n🔄 轉換數值欄位...")

for col in [total_col, subtotal_col, discount_col, shipping_col]:
    if col and col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# === 7. 計算統計數據 ===
print("\n💰 計算統計數據...")

# 總訂單數
total_orders = len(df)

# === 這裡是你要的：買家數量（不重複買家人數）===
# 你的資料有48行，買家不重複人數應該就是48人（如果每個買家只下一單）
# 但如果有人下多單，不重複人數會少於48

unique_buyers = df[buyer_col].dropna().nunique()
print(f"\n📊 買家統計：")
print(f"   - 總訂單數：{total_orders} 筆")
print(f"   - 不重複買家人數：{unique_buyers} 人")

# 如果總訂單數 = 48，不重複買家 = 48，表示每人只下一單
# 如果總訂單數 > 不重複買家，表示有人下多單

# 找出重複購買的買家
buyer_order_counts = df[buyer_col].value_counts()
repeat_buyers = buyer_order_counts[buyer_order_counts > 1].count()
one_time_buyers = unique_buyers - repeat_buyers

# 總金額
total_amount = df[total_col].sum() if total_col else 0

# 總小計
total_subtotal = df[subtotal_col].sum() if subtotal_col else 0

# 總折抵
total_discount = df[discount_col].sum() if discount_col else 0

# 總運費
total_shipping = df[shipping_col].sum() if shipping_col else 0

# 計算佔比
subtotal_percentage = (total_subtotal / total_amount * 100) if total_amount > 0 else 0
discount_percentage = (total_discount / total_amount * 100) if total_amount > 0 else 0
shipping_percentage = (total_shipping / total_amount * 100) if total_amount > 0 else 0

# 平均客單價
avg_order_value = total_amount / total_orders if total_orders > 0 else 0

# 平均每買家貢獻
avg_per_buyer = total_amount / unique_buyers if unique_buyers > 0 else 0

# 最高/最低單筆金額
max_amount = df[total_col].max() if total_col else 0
min_amount = df[total_col].min() if total_col else 0

# 有折抵的訂單
if discount_col:
    discount_orders = len(df[df[discount_col] > 0])
    discount_order_percentage = (discount_orders / total_orders * 100) if total_orders > 0 else 0
else:
    discount_orders = 0
    discount_order_percentage = 0

# 有運費的訂單
if shipping_col:
    shipping_orders = len(df[df[shipping_col] > 0])
    shipping_order_percentage = (shipping_orders / total_orders * 100) if total_orders > 0 else 0
else:
    shipping_orders = 0
    shipping_order_percentage = 0

# 重複購買分析
repeat_rate = (repeat_buyers / unique_buyers * 100) if unique_buyers > 0 else 0
avg_orders_per_buyer = total_orders / unique_buyers if unique_buyers > 0 else 0

print(f"\n📊 計算結果：")
print(f"   - 總訂單數：{total_orders}")
print(f"   - 買家數量（不重複）：{unique_buyers} 人")
print(f"   - 一次性買家：{one_time_buyers} 人")
print(f"   - 重複購買買家：{repeat_buyers} 人")
print(f"   - 重複購買率：{repeat_rate:.2f}%")
print(f"   - 平均每人下單次數：{avg_orders_per_buyer:.2f}")
print(f"   - 總金額：{total_amount:,.2f}")

# === 8. 建立統計表 ===
print("\n📋 建立統計報表...")

stats_data = {
    '統計項目': [
        '📦 訂單概況',
        '總訂單數 (筆)',
        '買家數量 (不重複人數)',
        '一次性買家人數',
        '重複購買買家人數',
        '重複購買率',
        '平均每人下單次數',
        '平均客單價',
        '平均每買家貢獻',
        '',
        '💰 金額分析',
        '總金額',
        '總小計 (商品金額)',
        '總折抵 (折扣/優惠)',
        '總運費',
        '',
        '📊 佔比分析',
        '小計佔總金額比例',
        '折抵佔總金額比例',
        '運費佔總金額比例',
        '',
        '📈 極值分析',
        '最高單筆金額',
        '最低單筆金額',
        '',
        '🏷️ 折抵分析',
        '有折抵的訂單數',
        '折抵訂單佔比',
        '',
        '🚚 運費分析',
        '有運費的訂單數',
        '運費訂單佔比'
    ],
    '數值': [
        '',
        f"{total_orders:,} 筆",
        f"{unique_buyers:,} 人",
        f"{one_time_buyers:,} 人",
        f"{repeat_buyers:,} 人",
        f"{repeat_rate:.2f}%",
        f"{avg_orders_per_buyer:.2f} 次",
        f"${avg_order_value:,.2f}",
        f"${avg_per_buyer:,.2f}",
        '',
        '',
        f"${total_amount:,.2f}",
        f"${total_subtotal:,.2f}",
        f"${total_discount:,.2f}",
        f"${total_shipping:,.2f}",
        '',
        '',
        f"{subtotal_percentage:.2f}%",
        f"{discount_percentage:.2f}%",
        f"{shipping_percentage:.2f}%",
        '',
        '',
        f"${max_amount:,.2f}",
        f"${min_amount:,.2f}",
        '',
        '',
        f"{discount_orders:,} 筆",
        f"{discount_order_percentage:.2f}%",
        '',
        '',
        f"{shipping_orders:,} 筆",
        f"{shipping_order_percentage:.2f}%"
    ]
}

stats_df = pd.DataFrame(stats_data)

# === 9. 儲存報表 ===
print(f"\n💾 正在更新統計表：{output_path}")

with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    stats_df.to_excel(writer, sheet_name='2025統計', index=False)

print(f"✅ 完成！已更新：{output_path}")

# === 10. 顯示摘要 ===
print("\n" + "=" * 60)
print("📊 Pinkoi 2025年統計摘要")
print("=" * 60)

print(f"\n📦 訂單概況：")
print(f"   ├─ 總訂單數：{total_orders:,} 筆")
print(f"   ├─ 買家數量：{unique_buyers:,} 人")  # 這是你要的48人！
print(f"   ├─ 一次性買家：{one_time_buyers:,} 人")
print(f"   ├─ 重複購買買家：{repeat_buyers:,} 人")
print(f"   ├─ 重複購買率：{repeat_rate:.2f}%")
print(f"   ├─ 平均每人下單：{avg_orders_per_buyer:.2f} 次")
print(f"   ├─ 平均客單價：${avg_order_value:,.2f}")
print(f"   └─ 平均每買家貢獻：${avg_per_buyer:,.2f}")

print("\n" + "=" * 60)
print("🎉 統計更新完成！")
print("=" * 60)


# In[ ]:




