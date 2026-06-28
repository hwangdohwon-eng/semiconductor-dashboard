"""
반도체 공정 데이터 분석 대시보드 v2
- OLS 회귀분석 제거
- 각 분석 파트 해석 가이드 강화
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.tri as tri
from scipy import stats
import statsmodels.api as sm
import seaborn as sns
import warnings
warnings.filterwarnings("ignore")

import platform
if platform.system() == "Windows":
    plt.rcParams["font.family"] = "Malgun Gothic"
elif platform.system() == "Darwin":
    plt.rcParams["font.family"] = "AppleGothic"
else:
    plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["axes.unicode_minus"] = False

st.set_page_config(page_title="반도체 공정 분석 대시보드", page_icon="🔬", layout="wide")

st.markdown("""
<style>
.main-header { text-align:center; padding:1.2rem 0 0.8rem; border-bottom:1px solid #e5e7eb; margin-bottom:1.5rem; }
.main-title  { font-size:26px; font-weight:700; color:#111827; }
.main-sub    { font-size:14px; color:#6b7280; margin-top:4px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
  <div class="main-title">🔬 반도체 공정 데이터 분석 대시보드</div>
  <div class="main-sub">엑셀 파일을 업로드하면 자동으로 공정 데이터를 분석합니다 | 취준생을 위한 현직자 수준 분석 툴</div>
</div>
""", unsafe_allow_html=True)

# ── 세션 초기화 ──────────────────────────────────────────────
def init():
    for k, v in {
        "df_eqp":None,"df_thk":None,"df_coor":None,"df_recipe":None,
        "run_cols":[],"eqp_num_cols":[],
    }.items():
        if k not in st.session_state:
            st.session_state[k] = v
init()

# ── 사이드바 ─────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📁 데이터 업로드")
    st.caption("파일을 업로드하면 자동으로 컬럼을 인식합니다.")
    f_thk    = st.file_uploader("📊 두께 데이터 (thickness)",          type=["xlsx","xls"], key="fu_thk")
    f_coor   = st.file_uploader("📍 좌표 데이터 (coordinate)",         type=["xlsx","xls"], key="fu_coor")
    f_eqp    = st.file_uploader("⚙️ 장비 파라미터 (eqp_parameters)",   type=["xlsx","xls"], key="fu_eqp")
    f_recipe = st.file_uploader("📋 레시피 데이터 (recipe) [선택]",    type=["xlsx","xls"], key="fu_rec")

    if st.button("🔄 데이터 로드 / 새로고침", type="primary", use_container_width=True):
        with st.spinner("데이터 불러오는 중..."):
            try:
                if f_thk:
                    xf = pd.ExcelFile(f_thk)
                    sh = "thickness" if "thickness" in xf.sheet_names else xf.sheet_names[0]
                    df_t = pd.read_excel(f_thk, sheet_name=sh)
                    df_t.columns = [str(c).strip() for c in df_t.columns]
                    run_cols = [c for c in df_t.columns if
                                any(k in str(c) for k in ["Run","run","Prerun","prerun"])]
                    for c in run_cols:
                        df_t[c] = pd.to_numeric(df_t[c], errors="coerce")
                    st.session_state.df_thk   = df_t
                    st.session_state.run_cols = run_cols

                if f_coor:
                    df_c = pd.read_excel(f_coor)
                    df_c.columns = [str(c).strip().lower() for c in df_c.columns]
                    st.session_state.df_coor = df_c

                if f_eqp:
                    df_e = pd.read_excel(f_eqp)
                    df_e.columns = [str(c).strip() for c in df_e.columns]
                    num_cols = df_e.select_dtypes(include=[np.number]).columns.tolist()
                    useful = [c for c in num_cols if
                              not any(x in c.lower() for x in ["id","no","number","count","index"])]
                    st.session_state.df_eqp       = df_e
                    st.session_state.eqp_num_cols = useful

                if f_recipe:
                    st.session_state.df_recipe = pd.read_excel(f_recipe)

                st.success("✅ 데이터 로드 완료!")
            except Exception as e:
                st.error(f"오류: {e}")

    st.divider()
    st.markdown("### 📌 분석 메뉴")
    menu = st.radio("", [
        "🏠 홈 — 데이터 개요",
        "📋 데이터 미리보기",
        "🗺️ 웨이퍼 맵 (Wafer Map)",
        "📈 분포 분석",
        "🔗 상관관계 분석",
        "📊 ANOVA",
        "📈 CPK 공정능력지수",
        "⏱️ 시계열 파라미터 추적",
        "🔵 불량 패턴 감지",
    ], label_visibility="collapsed")

    st.divider()
    if st.button("🗑️ 초기화", use_container_width=True):
        for k in ["df_eqp","df_thk","df_coor","df_recipe","run_cols","eqp_num_cols"]:
            st.session_state[k] = None if k not in ["run_cols","eqp_num_cols"] else []
        st.rerun()

# ── 공통 변수 ────────────────────────────────────────────────
df_thk       = st.session_state.df_thk
df_coor      = st.session_state.df_coor
df_eqp       = st.session_state.df_eqp
df_recipe    = st.session_state.df_recipe
run_cols     = st.session_state.run_cols
eqp_num_cols = st.session_state.eqp_num_cols

def need_thk():
    if df_thk is None or not run_cols:
        st.warning("⬅️ 먼저 두께 데이터를 업로드하고 [데이터 로드]를 눌러주세요.")
        return False
    return True

def need_coor():
    if df_coor is None:
        st.warning("⬅️ 좌표 데이터도 업로드해주세요.")
        return False
    return True

# ══════════════════════════════════════════════════════════════
# 🏠 홈
# ══════════════════════════════════════════════════════════════
if menu == "🏠 홈 — 데이터 개요":
    st.header("🏠 데이터 개요")
    if df_thk is None:
        st.info("👈 왼쪽에서 엑셀 파일을 업로드하고 **[데이터 로드]** 버튼을 눌러주세요.")
        st.markdown("""
        **업로드 가이드**
        | 업로드 칸 | 넣을 파일 | 필수 여부 |
        |-----------|-----------|-----------|
        | 📊 두께 데이터 | project2_thickness.xlsx | ✅ 필수 |
        | 📍 좌표 데이터 | project2_coordinate.xlsx | 웨이퍼 맵·불량패턴 시 필요 |
        | ⚙️ 장비 파라미터 | project2_eqp_paramters.xlsx | 상관분석·시계열 시 필요 |
        | 📋 레시피 | project2_dep_recipe.xlsx | 선택 |
        """)
        st.stop()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Run 수",          f"{len(run_cols)}개")
    col2.metric("측정 사이트 수",  f"{len(df_thk)}개")
    col3.metric("장비 파라미터 수",f"{len(eqp_num_cols)}개" if eqp_num_cols else "미업로드")
    all_vals = pd.concat([df_thk[c].dropna() for c in run_cols])
    col4.metric("전체 측정값",     f"{len(all_vals):,}개")

    st.divider()
    st.subheader("📊 Run별 평균 두께 추이")
    means = [df_thk[c].mean() for c in run_cols]
    stds  = [df_thk[c].std()  for c in run_cols]
    fig, ax = plt.subplots(figsize=(13, 4))
    ax.plot(run_cols, means, marker="o", color="steelblue", lw=2, label="평균 두께")
    ax.fill_between(range(len(run_cols)),
                    [m-s for m,s in zip(means,stds)],
                    [m+s for m,s in zip(means,stds)],
                    alpha=0.15, color="steelblue", label="±1σ 범위")
    ax.axhline(np.mean(means), color="red", linestyle="--",
               label=f"전체 평균: {np.mean(means):.1f} Å")
    ax.set_xticks(range(len(run_cols)))
    ax.set_xticklabels(run_cols, rotation=45, ha="right")
    ax.set_ylabel("두께 (Å)"); ax.set_title("Run별 평균 두께 추이")
    ax.legend(); ax.grid(True, alpha=0.3); plt.tight_layout()
    st.pyplot(fig); plt.close()

    st.markdown("---")
    st.markdown("### 💡 홈 해석 가이드")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.info("📈 **드리프트(Drift) 확인**")
        st.markdown("Run이 진행될수록 평균 두께가 지속적으로 올라가거나 내려가면 **장비 드리프트** 신호입니다. 챔버 세정이나 파라미터 재조정이 필요합니다.")
    with col_b:
        st.warning("📊 **±1σ 범위 확인**")
        st.markdown("범위가 갑자기 넓어지는 Run은 **공정 불안정** 신호입니다. 해당 Run의 웨이퍼 맵을 추가로 확인하세요.")
    with col_c:
        st.success("📐 **균일도(%) 기준**")
        st.markdown("균일도는 **낮을수록 좋습니다.** 보통 산업 현장에서는 **1% 이하**를 목표로 하며, 2% 초과 시 공정 조건 점검이 필요합니다.")

    st.subheader("📋 Run별 통계 요약")
    rows = []
    for c in run_cols:
        v = df_thk[c].dropna()
        rows.append({"Run":c, "평균(Å)":round(v.mean(),1),
                     "표준편차":round(v.std(),2),
                     "최소":round(v.min(),1), "최대":round(v.max(),1),
                     "범위":round(v.max()-v.min(),1),
                     "균일도(%)":round(v.std()/v.mean()*100,3)})
    st.dataframe(pd.DataFrame(rows), use_container_width=True)

# ══════════════════════════════════════════════════════════════
# 📋 데이터 미리보기
# ══════════════════════════════════════════════════════════════
elif menu == "📋 데이터 미리보기":
    st.header("📋 데이터 미리보기")
    tabs, names = [], []
    if df_thk    is not None: tabs.append(df_thk);    names.append("두께 데이터")
    if df_coor   is not None: tabs.append(df_coor);   names.append("좌표 데이터")
    if df_eqp    is not None: tabs.append(df_eqp);    names.append("장비 파라미터")
    if df_recipe is not None: tabs.append(df_recipe); names.append("레시피")
    if not tabs:
        st.warning("업로드된 데이터가 없습니다.")
    else:
        for tab, df in zip(st.tabs(names), tabs):
            with tab:
                st.write(f"Shape: {df.shape[0]}행 × {df.shape[1]}열")
                st.dataframe(df, use_container_width=True)

# ══════════════════════════════════════════════════════════════
# 🗺️ 웨이퍼 맵
# ══════════════════════════════════════════════════════════════
elif menu == "🗺️ 웨이퍼 맵 (Wafer Map)":
    st.header("🗺️ 웨이퍼 맵 (Wafer Map)")
    if not need_thk() or not need_coor(): st.stop()

    x_col = next((c for c in df_coor.columns if "x" in c.lower()), None)
    y_col = next((c for c in df_coor.columns if "y" in c.lower()), None)
    if not x_col or not y_col:
        st.error("좌표 파일에서 x, y 컬럼을 찾을 수 없습니다."); st.stop()

    col1, col2 = st.columns([1, 3])
    with col1:
        selected_run = st.selectbox("Run 선택", run_cols)
        show_contour = st.checkbox("등고선 표시", value=True)
        show_labels  = st.checkbox("등고선 라벨", value=True)
        colormap     = st.selectbox("컬러맵", ["jet","viridis","plasma","coolwarm","RdYlBu_r"])

    thk_vals = df_thk[selected_run].values
    n = min(len(df_coor), len(thk_vals))
    df_plot = df_coor.iloc[:n].copy()
    df_plot["thickness"] = thk_vals[:n]
    df_plot = df_plot.dropna(subset=[x_col, y_col, "thickness"])

    x = df_plot[x_col].values
    y = df_plot[y_col].values
    z = df_plot["thickness"].values
    xc, yc = np.mean(x), np.mean(y)
    wafer_r = np.sqrt((x-xc)**2+(y-yc)**2).max()

    with col2:
        triang = tri.Triangulation(x, y)
        tx = x[triang.triangles].mean(axis=1)
        ty = y[triang.triangles].mean(axis=1)
        triang.set_mask(np.sqrt((tx-xc)**2+(ty-yc)**2) > wafer_r*1.01)

        vmin = pd.concat([df_thk[c].dropna() for c in run_cols]).min()
        vmax = pd.concat([df_thk[c].dropna() for c in run_cols]).max()

        fig, ax = plt.subplots(figsize=(5,5), dpi=110)
        filled = ax.tricontourf(triang, z, levels=np.linspace(vmin,vmax,200),
                                cmap=colormap, vmin=vmin, vmax=vmax)
        if show_contour:
            cl = ax.tricontour(triang, z, levels=np.linspace(vmin,vmax,15),
                               colors="black", linewidths=0.7)
            if show_labels:
                ax.clabel(cl, inline=True, fontsize=6, fmt="%.0f")
        ax.add_patch(plt.Circle((xc,yc), wafer_r, fill=False, color="black", lw=1.5))
        cb = plt.colorbar(filled, ax=ax, fraction=0.046, pad=0.04)
        cb.set_label("두께 (Å)", fontsize=8)
        ax.set_title(f"Wafer Map — {selected_run}", fontsize=10)
        ax.set_aspect("equal")
        ax.set_xlim(xc-wafer_r*1.08, xc+wafer_r*1.08)
        ax.set_ylim(yc-wafer_r*1.08, yc+wafer_r*1.08)
        ax.axis("off"); plt.tight_layout()
        st.pyplot(fig, use_container_width=False); plt.close()

    m1,m2,m3,m4 = st.columns(4)
    m1.metric("평균 두께", f"{z.mean():.1f} Å")
    m2.metric("표준편차",  f"{z.std():.2f} Å")
    m3.metric("균일도",    f"{z.std()/z.mean()*100:.3f} %")
    m4.metric("범위",      f"{z.max()-z.min():.1f} Å")

    st.markdown("---")
    st.markdown("### 💡 웨이퍼 맵 해석 가이드")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.error("🔴 Center 핫스팟")
        st.markdown("중심부만 두껍거나 얇은 패턴\n\n→ **샤워헤드 막힘** 또는\n**플라즈마 밀도 불균일** 의심")
    with col_b:
        st.warning("🟡 Edge 불균일")
        st.markdown("가장자리가 두껍거나 얇은 패턴\n\n→ **가스 흐름 불균일** 또는\n**히터 온도 분포** 문제")
    with col_c:
        st.info("🔵 한쪽 치우침")
        st.markdown("좌우 또는 상하 비대칭 패턴\n\n→ **웨이퍼 기울어짐** 또는\n**가스 편류** 의심")
    st.success("✅ **이상적인 웨이퍼 맵** : 전체가 동일한 색(균일한 두께). 등고선 간격이 넓을수록 균일도가 높습니다.")

# ══════════════════════════════════════════════════════════════
# 📈 분포 분석
# ══════════════════════════════════════════════════════════════
elif menu == "📈 분포 분석":
    st.header("📈 분포 분석")
    if not need_thk(): st.stop()

    selected_run = st.selectbox("Run 선택", run_cols)
    vals = df_thk[selected_run].dropna()

    col1, col2 = st.columns(2)
    with col1:
        fig, ax = plt.subplots(figsize=(6,4))
        ax.hist(vals, bins=15, color="steelblue", edgecolor="black", alpha=0.8)
        ax.axvline(vals.mean(), color="red",    linestyle="--", label=f"평균: {vals.mean():.1f}")
        ax.axvline(vals.mean()+vals.std(), color="orange", linestyle=":", label="+1σ")
        ax.axvline(vals.mean()-vals.std(), color="orange", linestyle=":", label="-1σ")
        ax.set_xlabel("두께 (Å)"); ax.set_ylabel("빈도")
        ax.set_title(f"{selected_run} 두께 분포"); ax.legend()
        plt.tight_layout(); st.pyplot(fig); plt.close()

    with col2:
        fig, ax = plt.subplots(figsize=(6,4))
        stats.probplot(vals, dist="norm", plot=ax)
        ax.set_title("정규 확률도 (Q-Q Plot)")
        plt.tight_layout(); st.pyplot(fig); plt.close()

    st.subheader("📦 전체 Run 박스플롯")
    fig, ax = plt.subplots(figsize=(14,5))
    data = [df_thk[c].dropna().values for c in run_cols]
    bp = ax.boxplot(data, labels=run_cols, patch_artist=True)
    for patch in bp["boxes"]:
        patch.set_facecolor("lightsteelblue")
    ax.set_xlabel("Run"); ax.set_ylabel("두께 (Å)")
    ax.set_title("Run별 두께 분포 박스플롯")
    plt.xticks(rotation=45, ha="right"); plt.tight_layout()
    st.pyplot(fig); plt.close()

    st.markdown("---")
    st.markdown("### 💡 분포 분석 해석 가이드")
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**히스토그램 해석**")
        st.markdown(
            "- 종 모양(정규분포)이 이상적입니다\n"
            "- 좌우 비대칭(skewed)이면 공정 편향 의심\n"
            "- 두 개의 봉우리(bimodal)면 두 가지 공정 조건이 섞인 것\n"
            "- 분포가 넓을수록 균일도가 낮음"
        )
    with col_b:
        st.markdown("**Q-Q Plot 해석**")
        st.markdown(
            "- 점들이 직선에 가까울수록 정규분포\n"
            "- 양 끝이 직선에서 벗어나면 이상치 존재\n"
            "- S자 형태면 분포가 비대칭\n"
            "- **박스플롯** : 박스 크기가 클수록 산포(variability)가 큼"
        )
    stat_val, p_val = stats.shapiro(vals[:50] if len(vals)>50 else vals)
    if p_val > 0.05:
        st.success(f"📐 Shapiro-Wilk 정규성 검정 [{selected_run}] : W={stat_val:.4f}, p={p_val:.4f} → 정규분포로 볼 수 있음 ✅")
    else:
        st.warning(f"📐 Shapiro-Wilk 정규성 검정 [{selected_run}] : W={stat_val:.4f}, p={p_val:.4f} → 정규분포 아닐 수 있음 ⚠️ (이상치 또는 공정 변화 확인 필요)")

# ══════════════════════════════════════════════════════════════
# 🔗 상관관계 분석
# ══════════════════════════════════════════════════════════════
elif menu == "🔗 상관관계 분석":
    st.header("🔗 상관관계 분석")
    if not need_thk(): st.stop()

    tab1, tab2 = st.tabs(["Run 간 상관관계", "장비 파라미터 × 두께"])

    with tab1:
        df_runs = df_thk[run_cols].dropna()
        fig, ax = plt.subplots(figsize=(12,9))
        sns.heatmap(df_runs.corr(), annot=True, fmt=".2f", cmap="coolwarm",
                    vmin=-1, vmax=1, ax=ax, annot_kws={"size":8})
        ax.set_title("Run 간 두께 상관관계 히트맵")
        plt.tight_layout(); st.pyplot(fig); plt.close()

        st.markdown("---")
        st.markdown("### 💡 상관관계 해석 가이드")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.success("🟢 상관계수 ≥ 0.8")
            st.markdown("두 Run의 두께 패턴이 매우 유사 → 공정 조건이 안정적으로 유지되고 있음")
        with col_b:
            st.warning("🟡 0.5 ≤ 상관계수 < 0.8")
            st.markdown("어느 정도 유사하지만 차이 존재 → 공정 조건이 소폭 변화했을 가능성")
        with col_c:
            st.error("🔴 상관계수 < 0.5")
            st.markdown("두 Run의 패턴이 다름 → 공정 조건 변경 또는 장비 이상 의심. ANOVA로 추가 확인 필요")
        st.info("📌 **실무 팁** : 인접한 Run끼리 상관계수가 낮으면 해당 구간에서 공정 변화가 있었음을 의미합니다.")

    with tab2:
        if df_eqp is None or not eqp_num_cols:
            st.warning("장비 파라미터 데이터를 업로드해주세요.")
        else:
            target_run = st.selectbox("분석할 두께 Run", run_cols)
            sel_params = st.multiselect("분석할 파라미터 선택",
                                        eqp_num_cols, default=eqp_num_cols[:6])
            if sel_params:
                n = min(len(df_eqp), len(df_thk))
                df_merge = df_eqp[sel_params].iloc[:n].copy().reset_index(drop=True)
                df_merge["두께_mean"] = df_thk[target_run].values[:n]
                df_merge = df_merge.dropna()
                fig, ax = plt.subplots(figsize=(10,8))
                sns.heatmap(df_merge.corr(), annot=True, fmt=".2f",
                            cmap="coolwarm", vmin=-1, vmax=1, ax=ax)
                ax.set_title(f"장비 파라미터 × {target_run} 상관관계")
                plt.tight_layout(); st.pyplot(fig); plt.close()
                st.info("💡 두께와 상관계수 절댓값이 클수록 해당 파라미터가 두께에 큰 영향을 미칩니다. 양(+)의 상관 = 파라미터 증가 시 두께 증가, 음(-)의 상관 = 파라미터 증가 시 두께 감소.")

# ══════════════════════════════════════════════════════════════
# 📊 ANOVA
# ══════════════════════════════════════════════════════════════
elif menu == "📊 ANOVA":
    st.header("📊 분산분석 (ANOVA)")
    if not need_thk(): st.stop()

    st.markdown("**세 개 이상 Run 간 평균 두께 차이가 통계적으로 유의미한지 검정합니다.**")
    ref_run  = st.selectbox("기준 Run 선택", run_cols)
    ref_data = df_thk[ref_run].dropna()

    rows = []
    for c in run_cols:
        v = df_thk[c].dropna()
        f_val, p_val = stats.f_oneway(ref_data, v)
        rows.append({"Run":c, "F-통계량":round(f_val,4), "p-value":round(p_val,4),
                     "유의성":"✅ 유의 (p<0.05)" if p_val<0.05 else "❌ 유의하지 않음"})
    df_anova = pd.DataFrame(rows)
    st.dataframe(df_anova, use_container_width=True)

    fig, ax1 = plt.subplots(figsize=(13,5))
    df_s = df_anova.sort_values("F-통계량", ascending=False)
    colors = ["tomato" if p<0.05 else "steelblue" for p in df_s["p-value"]]
    ax1.bar(df_s["Run"], df_s["F-통계량"], color=colors, edgecolor="black", alpha=0.85, label="F-통계량")
    ax1.set_ylabel("F-통계량"); ax1.set_xlabel("Run")
    ax1.set_title(f"ANOVA F-통계량 / p-value (기준: {ref_run})")
    ax2 = ax1.twinx()
    ax2.plot(df_s["Run"], df_s["p-value"], color="orange", marker="o", lw=2, label="p-value")
    ax2.axhline(0.05, color="red", linestyle="--", lw=1.5, label="p=0.05 기준선")
    ax2.set_ylabel("p-value")
    l1,lb1 = ax1.get_legend_handles_labels()
    l2,lb2 = ax2.get_legend_handles_labels()
    ax1.legend(l1+l2, lb1+lb2, loc="upper right")
    plt.xticks(rotation=45, ha="right"); plt.tight_layout()
    st.pyplot(fig); plt.close()

    st.markdown("---")
    st.markdown("### 💡 ANOVA 해석 가이드")
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**F-통계량**")
        st.markdown(
            "- 값이 클수록 기준 Run과 두께 차이가 큼\n"
            "- 특정 Run에서 F값이 급격히 높아지면 해당 Run에서 **공정 변화 발생** 신호\n"
            "- 연속된 Run에서 F값이 증가 추세이면 **장비 드리프트(drift)** 의심"
        )
    with col_b:
        st.markdown("**p-value**")
        st.markdown(
            "- **p < 0.05 (빨간 막대)** : 기준 Run과 통계적으로 유의미한 차이 → 공정 이상 가능성\n"
            "- **p ≥ 0.05 (파란 막대)** : 기준 Run과 유사한 공정 상태\n"
            "- 빨간 막대가 많을수록 공정 안정성이 낮음을 의미"
        )
    st.info("📌 **실무 활용** : ANOVA에서 유의미한 Run을 발견했다면 해당 Run의 **웨이퍼 맵**을 확인해 이상 패턴을 시각적으로 검증하세요.")

# ══════════════════════════════════════════════════════════════
# 📈 CPK
# ══════════════════════════════════════════════════════════════
elif menu == "📈 CPK 공정능력지수":
    st.header("📈 CPK 공정능력지수")
    if not need_thk(): st.stop()

    col1, col2, col3 = st.columns(3)
    default_target = float(round(df_thk[run_cols].mean().mean()))
    default_std    = float(df_thk[run_cols[0]].std())
    with col1: target = st.number_input("목표 두께 (Target, Å)", value=default_target)
    with col2: usl    = st.number_input("USL (상한 규격)",       value=float(round(target+3*default_std)))
    with col3: lsl    = st.number_input("LSL (하한 규격)",       value=float(round(target-3*default_std)))

    rows = []
    for c in run_cols:
        v = df_thk[c].dropna()
        mu, sig = v.mean(), v.std()
        cp  = (usl-lsl)/(6*sig) if sig>0 else 0
        cpu = (usl-mu) /(3*sig) if sig>0 else 0
        cpl = (mu-lsl) /(3*sig) if sig>0 else 0
        cpk = min(cpu, cpl)
        rows.append({"Run":c,"평균":round(mu,1),"Std":round(sig,2),
                     "Cp":round(cp,3),"Cpk":round(cpk,3),
                     "판정":"🟢 우수" if cpk>=1.33 else "🟡 보통" if cpk>=1.0 else "🔴 불량"})
    df_cpk = pd.DataFrame(rows)
    st.dataframe(df_cpk, use_container_width=True)

    fig, ax = plt.subplots(figsize=(13,5))
    colors = ["green" if v>=1.33 else "orange" if v>=1.0 else "red" for v in df_cpk["Cpk"]]
    ax.bar(df_cpk["Run"], df_cpk["Cpk"], color=colors, edgecolor="black", alpha=0.85)
    ax.axhline(1.33, color="green",  linestyle="--", label="Cpk=1.33 (우수 기준)")
    ax.axhline(1.0,  color="orange", linestyle="--", label="Cpk=1.0  (최소 기준)")
    ax.set_xlabel("Run"); ax.set_ylabel("Cpk"); ax.set_title("Run별 Cpk")
    ax.legend(); plt.xticks(rotation=45, ha="right"); plt.tight_layout()
    st.pyplot(fig); plt.close()

    st.markdown("---")
    st.markdown("### 💡 CPK 해석 가이드")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.success("🟢 Cpk ≥ 1.33 : 우수")
        st.markdown("공정이 규격 내에서 충분한 여유를 두고 안정적으로 운영되고 있습니다. 현상 유지 권장.")
    with col_b:
        st.warning("🟡 1.0 ≤ Cpk < 1.33 : 주의")
        st.markdown("공정이 규격을 간신히 만족하는 수준. 공정 조건 점검 및 변동 원인 파악이 필요합니다.")
    with col_c:
        st.error("🔴 Cpk < 1.0 : 개선 필요")
        st.markdown("규격을 벗어나는 제품 발생 가능성 높음. 즉각적인 공정 파라미터 조정이 필요합니다.")
    st.info("📌 **Cp vs Cpk** : Cp는 공정 능력 자체, Cpk는 공정이 목표값에 얼마나 잘 맞는지를 반영합니다. **Cp는 높은데 Cpk가 낮으면** 공정이 한쪽으로 치우쳐(편향) 있다는 뜻입니다.")

# ══════════════════════════════════════════════════════════════
# ⏱️ 시계열
# ══════════════════════════════════════════════════════════════
elif menu == "⏱️ 시계열 파라미터 추적":
    st.header("⏱️ 시계열 파라미터 추적 (Set vs Actual)")
    if df_eqp is None:
        st.warning("장비 파라미터 데이터를 업로드해주세요."); st.stop()

    all_cols  = df_eqp.columns.tolist()
    set_cols  = [c for c in all_cols if any(k in c.lower() for k in ["set","setpoint"])]
    act_cols  = [c for c in all_cols if any(k in c.lower() for k in ["flow","actual","temp","pressure","power","forward"])]

    col1, col2 = st.columns(2)
    with col1: set_col = st.selectbox("Set 컬럼",    set_cols if set_cols else all_cols)
    with col2: act_col = st.selectbox("Actual 컬럼", act_cols if act_cols else all_cols)

    df_ts = df_eqp[[set_col, act_col]].copy()
    df_ts[set_col] = pd.to_numeric(df_ts[set_col].astype(str).str.extract(r"([\d.]+)")[0], errors="coerce")
    df_ts[act_col] = pd.to_numeric(df_ts[act_col].astype(str).str.extract(r"([\d.]+)")[0], errors="coerce")
    df_ts = df_ts.dropna().reset_index(drop=True)

    fig, ax = plt.subplots(figsize=(13,5))
    ax.plot(df_ts.index, df_ts[set_col], color="red", lw=2, label="Set value")
    ax.scatter(df_ts.index, df_ts[act_col], color="steelblue", s=20, alpha=0.6, label="Actual value")
    ax.set_xlabel("Data Index"); ax.set_ylabel("값")
    ax.set_title(f"Set vs Actual : {set_col} vs {act_col}")
    ax.legend(); ax.grid(True, alpha=0.3); plt.tight_layout()
    st.pyplot(fig); plt.close()

    diff = df_ts[act_col] - df_ts[set_col]
    c1,c2,c3 = st.columns(3)
    c1.metric("평균 오차", f"{diff.mean():.3f}")
    c2.metric("표준편차",  f"{diff.std():.3f}")
    c3.metric("최대 오차", f"{diff.abs().max():.3f}")

    st.markdown("---")
    st.markdown("### 💡 시계열 해석 가이드")
    col_a, col_b = st.columns(2)
    with col_a:
        st.success("✅ 정상 패턴")
        st.markdown(
            "- Actual(점)이 Set(빨간선)에 촘촘하게 붙어 있음\n"
            "- 오차가 일정하고 작음\n"
            "- 초반 안정화(settling) 후 일정하게 유지"
        )
    with col_b:
        st.error("⚠️ 이상 패턴")
        st.markdown(
            "- Actual이 Set에서 점점 멀어지는 추세 → **장비 드리프트**\n"
            "- 특정 구간에서 갑자기 튀는 값 → **공정 이상 이벤트**\n"
            "- 진동(oscillation) 패턴 → **제어 파라미터 불안정**"
        )
    st.info("📌 **실무 팁** : 평균 오차가 크거나 최대 오차가 평균의 3배 이상이면 해당 파라미터의 장비 캘리브레이션(calibration)을 점검해야 합니다.")

# ══════════════════════════════════════════════════════════════
# 🔵 불량 패턴 감지
# ══════════════════════════════════════════════════════════════
elif menu == "🔵 불량 패턴 감지":
    st.header("🔵 불량 패턴 감지")
    if not need_thk() or not need_coor(): st.stop()

    x_col = next((c for c in df_coor.columns if "x" in c.lower()), None)
    y_col = next((c for c in df_coor.columns if "y" in c.lower()), None)

    col1, col2 = st.columns([1,3])
    with col1:
        selected_run = st.selectbox("Run 선택", run_cols)
        threshold    = st.slider("불량 기준 (σ)", 1.0, 3.0, 2.0, 0.5)

    thk_vals = df_thk[selected_run].values
    n = min(len(df_coor), len(thk_vals))
    df_p = df_coor.iloc[:n].copy()
    df_p["thickness"] = thk_vals[:n]
    df_p = df_p.dropna(subset=[x_col, y_col, "thickness"])

    mu, sig = df_p["thickness"].mean(), df_p["thickness"].std()
    df_p["상태"] = df_p["thickness"].apply(
        lambda v: "불량(High)" if v>mu+threshold*sig
        else ("불량(Low)" if v<mu-threshold*sig else "정상"))

    color_map = {"정상":"green","불량(High)":"red","불량(Low)":"blue"}
    with col2:
        fig, ax = plt.subplots(figsize=(6,6))
        for status, grp in df_p.groupby("상태"):
            ax.scatter(grp[x_col], grp[y_col], c=color_map[status],
                       label=status, s=180, edgecolors="black", lw=0.5, zorder=3)
        xc, yc = df_p[x_col].mean(), df_p[y_col].mean()
        wr = np.sqrt((df_p[x_col]-xc)**2+(df_p[y_col]-yc)**2).max()
        ax.add_patch(plt.Circle((xc,yc), wr, color="gray", fill=False, linestyle="--"))
        ax.set_aspect("equal")
        ax.set_title(f"{selected_run} 불량 패턴 (±{threshold}σ 기준)")
        ax.set_xlabel("X (mm)"); ax.set_ylabel("Y (mm)")
        ax.legend(); ax.grid(True, alpha=0.2); plt.tight_layout()
        st.pyplot(fig); plt.close()

    good = (df_p["상태"]=="정상").sum()
    bad  = (df_p["상태"]!="정상").sum()
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("정상 사이트", f"{good}개")
    c2.metric("불량 사이트", f"{bad}개")
    c3.metric("불량률",      f"{bad/(good+bad)*100:.1f}%")
    c4.metric("기준",        f"±{threshold}σ")

    st.markdown("---")
    st.markdown("### 💡 불량 패턴 해석 가이드")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.error("🔴 가장자리(Edge) 집중")
        st.markdown(
            "웨이퍼 외곽에 불량 집중\n\n"
            "→ 가스 흐름 불균일\n"
            "→ 히터 온도 분포 문제\n"
            "→ Edge ring 마모 의심"
        )
    with col_b:
        st.error("🔴 중심부(Center) 집중")
        st.markdown(
            "웨이퍼 중심에 불량 집중\n\n"
            "→ 샤워헤드 막힘\n"
            "→ 플라즈마 밀도 불균일\n"
            "→ 가스 유량 조정 필요"
        )
    with col_c:
        st.error("🔴 한쪽 치우침(Asymmetry)")
        st.markdown(
            "불량이 한쪽으로 치우침\n\n"
            "→ 웨이퍼 기울어짐(tilt)\n"
            "→ 가스 편류(gas drift)\n"
            "→ 척(chuck) 수평 점검 필요"
        )
    st.info("📌 **분석 팁** : σ 기준을 **2.0 → 1.5**로 낮추면 더 민감하게 불량 후보를 감지할 수 있습니다. **3.0**으로 높이면 확실한 불량만 표시됩니다.")
