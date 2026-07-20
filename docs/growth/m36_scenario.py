#!/usr/bin/env python3
"""M36 目標逆算モデル: 月間営業利益 ¥83万 (年換算1,000万) の構成検証。

使い方:
    python3 docs/growth/m36_scenario.py

チャネル別のユニットエコノミクスを積み上げ、3シナリオ (保守/標準/強気) の
M36 月次PLを出力する。パラメータの根拠は 02-initiative-playbook.md を参照。
数値を変えて感度を見るのはこのファイルを直接編集すればよい
(月次18ヶ月モデルは docs/finance/build_financial_model.py、こちらは M36 の静的断面)。
"""

PAY = 0.0355          # Shopify Payments 決済手数料
LOSS = 0.015          # 破損・不良率
SHIP_COST = 999       # 発送実費/注文 (ゆうパック加重)
SHIP_INCOME = 880     # 顧客負担送料/注文 (無料ライン未到達)
MAT = 308             # 梱包資材/注文
FUL = 350             # 出荷外注 (3PL/アルバイト)/注文 — M18頃から発生


def own_ec(product_rev, aov, kake):
    """自社EC: (売上計, 貢献利益)。売上計 = 商品売上 + 送料収入。"""
    orders = product_rev / aov
    rev = product_rev + orders * SHIP_INCOME
    contrib = (product_rev * (1 - kake) + orders * SHIP_INCOME
               - rev * PAY - orders * (SHIP_COST + MAT + FUL)
               - product_rev * LOSS)
    return rev, contrib, orders


def mall(product_rev, aov, kake, fee):
    """モール: 送料込み価格・手数料は広告含む実質率。"""
    orders = product_rev / aov
    contrib = (product_rev * (1 - kake) - product_rev * fee
               - orders * (SHIP_COST + MAT + FUL) - product_rev * LOSS)
    return product_rev, contrib, orders


def tob(wholesale_rev, cost_rate_of_joudai=0.40, wholesale_rate=0.60, fee=0.05):
    """toB卸: 買取6掛。原価は上代比。出荷実費は売上比4%と粗く置く。"""
    joudai = wholesale_rev / wholesale_rate
    contrib = wholesale_rev - joudai * cost_rate_of_joudai \
        - wholesale_rev * fee - wholesale_rev * 0.04
    return wholesale_rev, contrib, None


SCENARIOS = {
    # (自社EC商品売上, AOV, 掛率) / モール系 / toB / その他
    "保守: セッション-30%・OEM4割止まり": {
        "自社EC (器+キッチン+食品)": own_ec(1_800_000, 6000, 0.55),
        "Yahoo!+キナリノ": mall(400_000, 6500, 0.55, 0.12),
        "eギフト (AnyGift/LINE)": mall(250_000, 6500, 0.52, 0.20),
        "toB卸 (テスト規模)": tob(300_000),
        "体験・ポップアップ": (100_000, 25_000, None),
        "_fixed": 145_000,
    },
    "標準: セッション1.6万・OEM6割": {
        "自社EC (器+キッチン+食品)": own_ec(2_520_000, 6000, 0.52),
        "Yahoo!+キナリノ": mall(500_000, 6500, 0.52, 0.12),
        "楽天 (ギフトセット型)": mall(800_000, 8500, 0.48, 0.24),
        "eギフト (AnyGift/LINE)": mall(300_000, 6500, 0.50, 0.20),
        "toB卸 (15〜20店)": tob(600_000),
        "体験・ポップアップ": (100_000, 30_000, None),
        "タイアップ記事 (月1本)": (80_000, 64_000, None),
        "_fixed": 173_000,
    },
    "強気: セッション2.2万・OEM7割": {
        "自社EC (器+キッチン+食品)": own_ec(3_300_000, 6200, 0.50),
        "Yahoo!+キナリノ": mall(600_000, 6500, 0.50, 0.12),
        "楽天 (ギフトセット型)": mall(1_200_000, 8500, 0.46, 0.24),
        "eギフト (AnyGift/LINE)": mall(400_000, 6500, 0.48, 0.20),
        "toB卸 (30店)": tob(900_000),
        "体験・ポップアップ": (150_000, 45_000, None),
        "タイアップ記事 (月2本)": (160_000, 128_000, None),
        "_fixed": 195_000,
    },
}

for name, ch in SCENARIOS.items():
    fixed = ch.pop("_fixed")
    rev_total = contrib_total = 0
    print(f"\n=== {name} ===")
    for label, (rev, contrib, orders) in ch.items():
        rev_total += rev
        contrib_total += contrib
        o = f" ({orders:.0f}注文)" if orders else ""
        print(f"  {label:<28} 月商 ¥{rev:>9,.0f}  貢献 ¥{contrib:>8,.0f} "
              f"({contrib / rev * 100:4.1f}%){o}")
    op = contrib_total - fixed
    print(f"  {'合計':<28} 月商 ¥{rev_total:>9,.0f}  貢献 ¥{contrib_total:>8,.0f} "
          f"({contrib_total / rev_total * 100:4.1f}%)")
    print(f"  固定費 ¥{fixed:,} → 営業利益 ¥{op:,.0f}/月 (年換算 ¥{op * 12 / 10_000:,.0f}万)")

# 自社EC 必要セッションの逆算
print("\n=== 自社EC 必要セッション (標準シナリオ) ===")
for label, product_rev, aov, cvr, repeat_share in [
    ("保守", 1_800_000, 6000, 0.015, 0.35),
    ("標準", 2_520_000, 6000, 0.016, 0.40),
    ("強気", 3_300_000, 6200, 0.018, 0.42),
]:
    orders = product_rev / aov
    sessions = orders * (1 - repeat_share) / cvr
    print(f"  {label}: 注文{orders:.0f}/月 × 新規比率{(1-repeat_share)*100:.0f}% ÷ CVR{cvr*100:.1f}%"
          f" = 必要セッション {sessions:,.0f}/月")
