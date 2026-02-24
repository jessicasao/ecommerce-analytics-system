#!/usr/bin/env python
# coding: utf-8

# In[1]:


# ============================================
# 程式名稱：整理 Shopify 訂單表
# 檔案路徑：C:\Users\MI\Desktop\2026-月度財務報表\01\202601-Shopify-Orders.xlsx
# 功能：
#   1. 欄位改名
#   2. 合併電話欄位
#   3. 加入 Index
# ============================================

import pandas as pd
import os
from datetime import datetime

print("=" * 60)
print("📦 開始整理 Shopify 訂單表...")
print("=" * 60)

# === 1. 設定檔案路徑 ===
input_path = r'C:\Users\MI\Desktop\2026-月度財務報表\01\202601-Shopify-Orders.xlsx'
output_path = r'C:\Users\MI\Desktop\2026-月度財務報表\01\202601-Shopify-Orders_整理版.xlsx'

# === 2. 檢查檔案是否存在 ===
if not os.path.exists(input_path):
    print(f"❌ 錯誤：找不到檔案")
    print(f"   路徑：{input_path}")
    exit()

# === 3. 讀取檔案 ===
print(f"\n📂 正在讀取：{input_path}")
df = pd.read_excel(input_path)
print(f"✅ 成功讀取：{len(df)} 行，{len(df.columns)} 欄")

# === 4. 顯示原始欄位 ===
print("\n📋 原始欄位：")
for i, col in enumerate(df.columns):
    print(f"  {i+1:2d}. '{col}'")

# === 5. 欄位改名對照表 ===
rename_map = {
    'Name': 'Order No',
    'Lineitem quantity': 'Quantity',
    'Lineitem name': 'Product Name',
    'Lineitem price': 'Cost',
    'Lineitem sku': 'SKU',
    'Billing Name': 'Customer Name',
    'Id': 'Order Id'
}

print("\n🔄 正在重新命名欄位...")
renamed_count = 0
for old_name, new_name in rename_map.items():
    if old_name in df.columns:
        df.rename(columns={old_name: new_name}, inplace=True)
        print(f"   ✅ '{old_name}' → '{new_name}'")
        renamed_count += 1
    else:
        print(f"   ⚠️ 找不到 '{old_name}'，跳過")

print(f"\n✅ 共重新命名 {renamed_count} 個欄位")

# === 6. 合併電話欄位 ===
print("\n📞 正在合併電話欄位...")

# 檢查電話欄位是否存在
has_billing_phone = 'Billing Phone' in df.columns
has_phone = 'Phone' in df.columns

if has_billing_phone or has_phone:
    # 確保兩個欄位都是字串類型
    if has_billing_phone:
        df['Billing Phone'] = df['Billing Phone'].astype(str).replace('nan', '').replace('None', '')
    if has_phone:
        df['Phone'] = df['Phone'].astype(str).replace('nan', '').replace('None', '')

    # 合併電話：優先使用 Billing Phone，如果沒有則用 Phone
    if has_billing_phone and has_phone:
        df['Phone'] = df.apply(
            lambda row: row['Billing Phone'] if row['Billing Phone'] and row['Billing Phone'].strip() 
            else (row['Phone'] if row['Phone'] and row['Phone'].strip() else ''),
            axis=1
        )
        print("   ✅ 已合併 Billing Phone 和 Phone → Phone")
        # 刪除 Billing Phone 欄位
        df.drop(columns=['Billing Phone'], inplace=True)
        print("   🗑️ 已刪除 Billing Phone 欄位")

    elif has_billing_phone and not has_phone:
        df.rename(columns={'Billing Phone': 'Phone'}, inplace=True)
        print("   ✅ Billing Phone 已改名為 Phone")

    # 處理空值
    df['Phone'] = df['Phone'].fillna('')
else:
    print("   ⚠️ 找不到任何電話欄位，新增空白欄位")
    df['Phone'] = ''

# === 7. 加入 Index 欄（放在第一欄） ===
print("\n🔢 正在加入 Index 欄位...")

# 建立 Index 欄位（從1開始）
df.insert(0, 'Index', range(1, len(df) + 1))
print(f"   ✅ 已加入 Index 欄 (1-{len(df)})")

# === 8. 顯示更新後的欄位 ===
print("\n📋 更新後的欄位：")
for i, col in enumerate(df.columns):
    print(f"  {i+1:2d}. '{col}'")

# === 9. 資料統計 ===
print("\n📊 資料統計：")
print(f"   - 總筆數：{len(df)} 筆")
print(f"   - 總欄位數：{len(df.columns)} 個")
print(f"   - 有電話的訂單：{(df['Phone'] != '').sum()} 筆")

# 如果有 Order No 欄位，顯示訂單範圍
if 'Order No' in df.columns:
    order_count = df['Order No'].nunique()
    print(f"   - 不重複訂單編號：{order_count} 個")

# === 10. 檢查是否有數值欄位 ===
numeric_cols = ['Quantity', 'Cost', 'Total']  # 假設有這些欄位
for col in numeric_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        print(f"   - {col} 總和：{df[col].sum():,.2f}")

# === 11. 儲存檔案 ===
print(f"\n💾 正在儲存檔案：{output_path}")
try:
    df.to_excel(output_path, index=False)
    print(f"✅ 完成！已儲存為：{output_path}")
except Exception as e:
    print(f"❌ 儲存失敗：{e}")
    # 嘗試用不同引擎
    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"✅ 使用 openpyxl 引擎儲存成功")

# === 12. 顯示前5筆資料 ===
print("\n👀 前5筆資料（主要欄位）：")
preview_cols = ['Index', 'Order No', 'Customer Name', 'Product Name', 'Quantity', 'Phone']
preview_cols = [col for col in preview_cols if col in df.columns]
print(df[preview_cols].head())

# === 13. 生成簡單的統計報表 ===
print("\n📈 簡易統計：")
if 'Order No' in df.columns and 'Quantity' in df.columns and 'Cost' in df.columns:
    total_orders = df['Order No'].nunique()
    total_quantity = df['Quantity'].sum()
    total_revenue = (df['Quantity'] * df['Cost']).sum()

    print(f"   總訂單數：{total_orders} 筆")
    print(f"   總銷售數量：{total_quantity:.0f} 件")
    print(f"   總營業額：${total_revenue:,.2f}")

print("\n" + "=" * 60)
print("🎉 整理完成！")
print("=" * 60)


# In[ ]:


DEEPSEEK指令

"C:\Users\MI\Desktop\2026-月度財務報表\01\202601-Shopify-Orders.xlsx" 把這個檔改以下, 用PYTHON

1.請保留以下欄位
Name	Email	Paid at	Accepts Marketing	Total	Discount Code	Discount Amount	Created at	Lineitem quantity	Lineitem name	Lineitem price	Lineitem sku		Billing Name	Billing Phone	Id	Source	Phone

"C:\Users\MI\Desktop\2026-月度財務報表\01\202601-Shopify-Orders.xlsx" 把這個檔改以下, 用PYTHON

2.欄位要改名:
Name 改Order No , Lineitem quantity 改  Quantity, Lineitem name 改 Product Name,
Lineitem price 改 Selling Price,  Lineitem sku 改Variant SKU, Billing Name改Customer Name , Id 改成 Order Id 

3.把Billing Phone和Phone 合拼成一欄, 叫Phone

4.加Index在第一欄, 日後方便數單數用





# In[ ]:





# In[ ]:


DeepSEEK
第二
把"C:\Users\MI\Desktop\2026-月度財務報表\01\Cost_with_ID_最終版.xlsx" 裡的Variant SKU 和 Cost 
加進 C:\Users\MI\Desktop\2026-月度財務報表\01\202601-Shopify-Orders_整理版.xlsx

用Product Name來連那二個表
成功后存在C:\Users\MI\Desktop\2026-月度財務報表\01\  
命名為Shopify-Orders_計算版


# In[2]:


# ============================================
# 程式名稱：用產品名稱加入成本和 SKU
# 訂單表：C:\Users\MI\Desktop\2026-月度財務報表\01\202601-Shopify-Orders_整理版.xlsx
# 成本表：C:\Users\MI\Desktop\2026-月度財務報表\01\Cost_with_ID_最終版.xlsx
# 共同欄位：Product Name
# 輸出：Shopify-Orders_計算版.xlsx
# ============================================

import pandas as pd
import os

print("=" * 60)
print("📦 開始用產品名稱加入成本和 SKU...")
print("=" * 60)

# === 1. 設定檔案路徑 ===
folder_path = r'C:\Users\MI\Desktop\2026-月度財務報表\01'
orders_path = os.path.join(folder_path, '202601-Shopify-Orders_整理版.xlsx')
cost_path = os.path.join(folder_path, 'Cost_with_ID_最終版.xlsx')
output_path = os.path.join(folder_path, 'Shopify-Orders_計算版.xlsx')

# === 2. 檢查檔案是否存在 ===
if not os.path.exists(orders_path):
    print(f"❌ 錯誤：找不到訂單表")
    print(f"   路徑：{orders_path}")
    exit()

if not os.path.exists(cost_path):
    print(f"❌ 錯誤：找不到成本表")
    print(f"   路徑：{cost_path}")
    exit()

# === 3. 讀取檔案 ===
print(f"\n📂 正在讀取訂單表：{orders_path}")
orders = pd.read_excel(orders_path)
print(f"✅ 訂單表：{len(orders)} 行，{len(orders.columns)} 欄")

print(f"\n📂 正在讀取成本表：{cost_path}")
cost = pd.read_excel(cost_path)
print(f"✅ 成本表：{len(cost)} 行，{len(cost.columns)} 欄")

# === 4. 顯示兩個表的欄位 ===
print("\n📋 訂單表欄位：")
for i, col in enumerate(orders.columns):
    print(f"  {i+1:2d}. '{col}'")

print("\n📋 成本表欄位：")
for i, col in enumerate(cost.columns):
    print(f"  {i+1:2d}. '{col}'")

# === 5. 確認產品名稱欄位是否存在 ===
# 訂單表的產品名稱欄位
order_product_col = None
possible_order_product_cols = ['Product Name', '產品名稱', 'Lineitem name', '商品名稱']
for col in possible_order_product_cols:
    if col in orders.columns:
        order_product_col = col
        break

if not order_product_col:
    print("\n❌ 錯誤：訂單表找不到產品名稱欄位")
    print("請確認訂單表是否有以下欄位之一：")
    for col in possible_order_product_cols:
        print(f"   - {col}")
    exit()

print(f"\n✅ 找到訂單表產品名稱欄位：『{order_product_col}』")

# 成本表的產品名稱欄位
cost_product_col = None
possible_cost_product_cols = ['Product_Name', '產品名稱', 'Product Name', '商品名稱']
for col in possible_cost_product_cols:
    if col in cost.columns:
        cost_product_col = col
        break

if not cost_product_col:
    print("\n❌ 錯誤：成本表找不到產品名稱欄位")
    print("請確認成本表是否有以下欄位之一：")
    for col in possible_cost_product_cols:
        print(f"   - {col}")
    exit()

print(f"✅ 找到成本表產品名稱欄位：『{cost_product_col}』")

# 確認成本表有需要的欄位
if 'Variant SKU' not in cost.columns:
    print("\n❌ 錯誤：成本表沒有『Variant SKU』欄位")
    exit()

if 'Cost' not in cost.columns:
    print("\n❌ 錯誤：成本表沒有『Cost』欄位")
    exit()

# === 6. 清理成本表資料 ===
print("\n🧹 清理成本表資料...")

# 移除產品名稱為空的行
cost_clean = cost[cost[cost_product_col].notna()].copy()

# 處理可能的重複產品名稱（保留第一個）
cost_clean = cost_clean.drop_duplicates(subset=[cost_product_col])

# 建立對照表
product_to_sku = dict(zip(cost_clean[cost_product_col], cost_clean['Variant SKU']))
product_to_cost = dict(zip(cost_clean[cost_product_col], cost_clean['Cost']))

print(f"✅ 共建立 {len(product_to_sku)} 個產品的對照")

# === 7. 把成本和 SKU 加到訂單表 ===
print(f"\n🔄 正在用『{order_product_col}』加入成本和 SKU...")

# 新增欄位
orders['Variant SKU'] = orders[order_product_col].map(product_to_sku)
orders['單位成本'] = orders[order_product_col].map(product_to_cost)

# 統計找到的情況
found_sku = orders['Variant SKU'].notna().sum()
found_cost = orders['單位成本'].notna().sum()
total_rows = len(orders)

print(f"\n📊 匹配結果：")
print(f"   - 總筆數：{total_rows}")
print(f"   - 找到 SKU：{found_sku} 筆 ({found_sku/total_rows*100:.1f}%)")
print(f"   - 找到成本：{found_cost} 筆 ({found_cost/total_rows*100:.1f}%)")

# 如果找不到，補空值
orders['Variant SKU'] = orders['Variant SKU'].fillna('')
orders['單位成本'] = orders['單位成本'].fillna(0)

# === 8. 找出找不到成本的產品 ===
if total_rows - found_cost > 0:
    print("\n⚠️ 找不到成本的產品：")
    missing_products = orders[orders['單位成本'] == 0][order_product_col].unique()
    for product in missing_products[:20]:  # 只顯示前20個
        print(f"   - {product}")
    if len(missing_products) > 20:
        print(f"   ... 還有 {len(missing_products) - 20} 個")

# === 9. 計算總成本和利潤 ===
print("\n💰 計算成本和利潤...")

# 確保數值格式
if 'Quantity' in orders.columns:
    orders['Quantity'] = pd.to_numeric(orders['Quantity'], errors='coerce').fillna(0)
else:
    print("⚠️ 找不到 Quantity 欄位，使用預設值 1")
    orders['Quantity'] = 1

if 'Cost' in orders.columns:
    orders['Cost'] = pd.to_numeric(orders['Cost'], errors='coerce').fillna(0)
else:
    print("⚠️ 找不到 Cost 欄位（單價），使用預設值 0")
    orders['Cost'] = 0

# 計算
orders['總成本'] = orders['Quantity'] * orders['單位成本']
orders['總售價'] = orders['Quantity'] * orders['Cost']
orders['利潤'] = orders['總售價'] - orders['總成本']
orders['毛利率'] = (orders['利潤'] / orders['總售價'] * 100).round(1)
orders.loc[orders['總售價'] == 0, '毛利率'] = 0

print(f"\n📈 總計：")
print(f"   總售價：{orders['總售價'].sum():,.2f}")
print(f"   總成本：{orders['總成本'].sum():,.2f}")
print(f"   總利潤：{orders['利潤'].sum():,.2f}")
if orders['總售價'].sum() > 0:
    print(f"   平均毛利率：{(orders['利潤'].sum() / orders['總售價'].sum() * 100):.1f}%")

# === 10. 調整欄位順序 ===
print("\n📋 調整欄位順序...")

# 找出 Product Name 的位置
columns = orders.columns.tolist()
if order_product_col in columns:
    name_idx = columns.index(order_product_col)

    # 重新排列：把新欄位放在 Product Name 旁邊
    new_order = []
    for col in columns:
        new_order.append(col)
        if col == order_product_col:
            new_order.append('Variant SKU')
            new_order.append('單位成本')
            new_order.append('總成本')
            new_order.append('總售價')
            new_order.append('利潤')
            new_order.append('毛利率')

    # 移除重複
    new_order = list(dict.fromkeys(new_order))
    orders = orders[new_order]
    print("✅ 欄位順序調整完成")

# === 11. 顯示更新後的欄位 ===
print("\n📋 更新後的欄位：")
for i, col in enumerate(orders.columns):
    print(f"  {i+1:2d}. '{col}'")

# === 12. 儲存檔案 ===
print(f"\n💾 正在儲存檔案：{output_path}")
orders.to_excel(output_path, index=False)
print(f"✅ 完成！已儲存為：{output_path}")

# === 13. 顯示前5筆資料 ===
print("\n👀 前5筆資料（主要欄位）：")
preview_cols = ['Index', 'Order No', order_product_col, 'Variant SKU', '單位成本', 'Quantity', 'Cost', '利潤', '毛利率']
preview_cols = [col for col in preview_cols if col in orders.columns]
print(orders[preview_cols].head())

# === 14. 產生簡易報表 ===
print("\n📊 簡易報表：")
print(f"   總訂單明細數：{total_rows} 筆")
print(f"   有成本的產品數：{found_cost} 筆")
print(f"   無成本的產品數：{total_rows - found_cost} 筆")
print(f"   總營業額：${orders['總售價'].sum():,.2f}")
print(f"   總成本：${orders['總成本'].sum():,.2f}")
print(f"   總利潤：${orders['利潤'].sum():,.2f}")

print("\n" + "=" * 60)
print("🎉 完成！")
print("=" * 60)


# In[ ]:





# In[ ]:


"C:\Users\MI\Desktop\2026-月度財務報表\01\Shopify-Orders_計算版.xlsx"把這個檔改以下, 用PYTHON

1.在金Cost 欄位隔離增加Profit 

2.再替我增加3欄, 計算: 總Profit , 總成本, 總Profit %

最後命名Shopify-Orders_計算版-V2.xlsx


# In[10]:


# ============================================
# 程式名稱：Shopify 訂單表 Total 和折扣分攤
# 檔案路徑：C:\Users\MI\Desktop\2026-月度財務報表\01\Shopify-Orders_計算版-V2.xlsx
# 功能：
#   1. 按商品金額比例分攤 Total 到每個商品
#   2. 按相同比例分攤 Discount Amount 到每個商品
#   3. 新增「分攤後金額」放在 Total 前面
# 輸出：Shopify-Orders_計算版-V3.xlsx
# ============================================

import pandas as pd
import os

print("=" * 60)
print("📦 開始分攤 Total 和 Discount Amount...")
print("=" * 60)

# === 1. 設定檔案路徑 ===
folder_path = r'C:\Users\MI\Desktop\2026-月度財務報表\01'
input_path = os.path.join(folder_path, 'Shopify-Orders_計算版-V2.xlsx')
output_path = os.path.join(folder_path, 'Shopify-Orders_計算版-V3.xlsx')

# === 2. 檢查檔案是否存在 ===
if not os.path.exists(input_path):
    print(f"❌ 錯誤：找不到檔案")
    print(f"   路徑：{input_path}")
    exit()

# === 3. 讀取檔案 ===
print(f"\n📂 正在讀取：{input_path}")
df = pd.read_excel(input_path)
print(f"✅ 成功讀取：{len(df)} 行，{len(df.columns)} 欄")

# === 4. 顯示原始欄位 ===
print("\n📋 原始欄位：")
for i, col in enumerate(df.columns):
    print(f"  {i+1:2d}. '{col}'")

# === 5. 確認必要的欄位存在 ===
required_cols = ['Order No', 'Selling Price', 'Quantity', 'Total', 'Discount Amount']
missing_cols = [col for col in required_cols if col not in df.columns]

if missing_cols:
    print(f"\n❌ 錯誤：缺少以下必要欄位：{missing_cols}")
    exit()

# === 6. 確保數值欄位格式正確 ===
print("\n🔄 確保數值欄位格式正確...")

df['Selling Price'] = pd.to_numeric(df['Selling Price'], errors='coerce').fillna(0)
df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
df['Total'] = pd.to_numeric(df['Total'], errors='coerce').fillna(0)
df['Discount Amount'] = pd.to_numeric(df['Discount Amount'], errors='coerce').fillna(0)

# === 7. 計算每個商品的原始金額 ===
df['商品金額'] = df['Selling Price'] * df['Quantity']
print(f"\n💰 商品金額範圍：{df['商品金額'].min():.2f} ~ {df['商品金額'].max():.2f}")

# === 8. 顯示原始資料 ===
print("\n📋 原始資料（前10行）：")
print(df[['Order No', 'Selling Price', 'Quantity', '商品金額', 'Total', 'Discount Amount']].head(10))

# === 9. 新增分攤欄位 ===
df['分攤後金額'] = 0.0
df['分攤後折扣'] = 0.0

# === 10. 按訂單分組進行分攤 ===
print("\n🔄 正在按比例分攤 Total 和 Discount...")

total_orders = df['Order No'].nunique()
processed = 0
problem_orders = []

for order_no, group in df.groupby('Order No'):
    processed += 1
    print(f"\r   處理中：{processed}/{total_orders} 筆訂單", end="")

    # 找出這個訂單的 Total（非0的那行）
    total_rows = group[group['Total'] > 0]
    if len(total_rows) > 0:
        order_total = total_rows['Total'].iloc[0]
    else:
        order_total = 0
        problem_orders.append(f"{order_no} (無 Total)")
        continue

    # 找出這個訂單的 Discount（非0的那行）
    discount_rows = group[group['Discount Amount'] > 0]
    if len(discount_rows) > 0:
        order_discount = discount_rows['Discount Amount'].iloc[0]
    else:
        order_discount = 0

    # 計算這個訂單所有商品的商品金額總和
    group_total_goods = group['商品金額'].sum()

    if group_total_goods > 0:
        # 按比例分攤到每個商品
        for idx in group.index:
            ratio = group.loc[idx, '商品金額'] / group_total_goods

            # 分攤 Total
            df.loc[idx, '分攤後金額'] = round(order_total * ratio, 2)

            # 分攤 Discount
            df.loc[idx, '分攤後折扣'] = round(order_discount * ratio, 2)
    else:
        problem_orders.append(f"{order_no} (商品金額為0)")

print("\n\n✅ 分攤完成")

# === 11. 顯示有問題的訂單 ===
if problem_orders:
    print(f"\n⚠️ 發現 {len(problem_orders)} 筆有問題的訂單：")
    for order in problem_orders[:10]:
        print(f"   - {order}")
    if len(problem_orders) > 10:
        print(f"   ... 還有 {len(problem_orders) - 10} 筆")

# === 12. 驗證分攤是否正確 ===
print("\n🔍 驗證分攤結果：")

verification = []
all_correct = True

for order_no, group in df.groupby('Order No'):
    # 原始 Total
    original_total = group[group['Total'] > 0]['Total'].iloc[0] if any(group['Total'] > 0) else 0

    # 原始 Discount
    original_discount = group[group['Discount Amount'] > 0]['Discount Amount'].iloc[0] if any(group['Discount Amount'] > 0) else 0

    # 分攤後 Total 加總
    allocated_total = group['分攤後金額'].sum()

    # 分攤後 Discount 加總
    allocated_discount = group['分攤後折扣'].sum()

    # 計算差異
    total_diff = abs(original_total - allocated_total)
    discount_diff = abs(original_discount - allocated_discount)

    is_correct = total_diff < 0.1 and discount_diff < 0.1

    verification.append({
        '訂單編號': order_no,
        '原始Total': original_total,
        '分攤後Total總和': allocated_total,
        'Total差異': total_diff,
        '原始Discount': original_discount,
        '分攤後Discount總和': allocated_discount,
        'Discount差異': discount_diff,
        '正確': '✅' if is_correct else '❌'
    })

    if not is_correct and original_total > 0:
        all_correct = False
        print(f"\n   ⚠️ 訂單 {order_no}：")
        print(f"     Total: 原始 {original_total:.2f} vs 分攤後 {allocated_total:.2f} (差異 {total_diff:.2f})")
        print(f"     Discount: 原始 {original_discount:.2f} vs 分攤後 {allocated_discount:.2f} (差異 {discount_diff:.2f})")

if all_correct:
    print("   ✅ 所有訂單分攤正確！")

# 建立驗證表格
verification_df = pd.DataFrame(verification)

# === 13. 調整欄位順序（把分攤後金額放在 Total 前面）===
print("\n📋 調整欄位順序...")

cols = df.columns.tolist()

# 找到 Total 的位置
if 'Total' in cols:
    total_idx = cols.index('Total')

    # 移除要移動的欄位
    if '分攤後金額' in cols:
        cols.remove('分攤後金額')
    if '分攤後折扣' in cols:
        cols.remove('分攤後折扣')
    if '商品金額' in cols:
        cols.remove('商品金額')

    # 在 Total 前面插入分攤後金額
    new_cols = cols[:total_idx] + ['分攤後金額'] + cols[total_idx:]

    # 在 Discount Amount 前面或後面插入分攤後折扣
    if 'Discount Amount' in new_cols:
        discount_idx = new_cols.index('Discount Amount')
        new_cols = new_cols[:discount_idx+1] + ['分攤後折扣'] + new_cols[discount_idx+1:]
    else:
        new_cols.append('分攤後折扣')

    # 把商品金額放在 Selling Price 旁邊
    if 'Selling Price' in new_cols:
        price_idx = new_cols.index('Selling Price')
        new_cols = new_cols[:price_idx+1] + ['商品金額'] + new_cols[price_idx+1:]

    df = df[new_cols]
    print("✅ 欄位順序調整完成")

# === 14. 顯示分攤結果範例 ===
print("\n📊 分攤結果範例（前10行）：")
result_cols = ['Order No', 'Selling Price', '商品金額', 'Quantity', 
               '分攤後金額', 'Total', '分攤後折扣', 'Discount Amount']
result_cols = [col for col in result_cols if col in df.columns]
print(df[result_cols].head(10))

# === 15. 計算總計 ===
print("\n📈 總計比較：")
total_before = df[df['Total'] > 0]['Total'].sum()
total_after = df['分攤後金額'].sum()
discount_before = df[df['Discount Amount'] > 0]['Discount Amount'].sum()
discount_after = df['分攤後折扣'].sum()

print(f"   Total 分攤前：{total_before:,.2f}")
print(f"   Total 分攤後：{total_after:,.2f}")
print(f"   差異：{total_after - total_before:,.2f}")
print(f"   Discount 分攤前：{discount_before:,.2f}")
print(f"   Discount 分攤後：{discount_after:,.2f}")
print(f"   差異：{discount_after - discount_before:,.2f}")

# === 16. 儲存檔案 ===
print(f"\n💾 正在儲存檔案：{output_path}")
df.to_excel(output_path, index=False)
print(f"✅ 完成！已儲存為：{output_path}")

# === 17. 儲存驗證結果 ===
verification_output = os.path.join(folder_path, '分攤驗證結果.xlsx')
verification_df.to_excel(verification_output, index=False)
print(f"✅ 驗證結果已儲存：{verification_output}")

print("\n" + "=" * 60)
print("🎉 完成！")
print("=" * 60)


# In[ ]:





# In[ ]:





# In[ ]:




