import os
import pandas as pd
import numpy as np
import json
import re
import warnings
from scipy.stats import skew, kurtosis, f_oneway, ttest_ind
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from sklearn.linear_model import LinearRegression
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler

# Suppress warnings for clean execution
warnings.filterwarnings('ignore')
os.environ['LOKY_MAX_CPU_COUNT'] = '1'

EXCEL_FILE = r"C:\Users\HP\Downloads\Marketing\data.xlsx"
HTML_FILE = "dash.html"

# NIRF Sub-metrics mapping
SUB_METRICS = ['SS', 'FSR', 'FQE', 'FRU', 'PU', 'QP', 'IPR', 'FPPP', 'GPH', 'GUE', 'MS', 'GPHD', 'RD', 'WD', 'ESCS', 'PCS', 'PR']

def clean_json(obj):
    """Recursively replace NaN/Inf with None for JSON compatibility."""
    if isinstance(obj, float):
        if np.isnan(obj) or np.isinf(obj):
            return None
        return round(obj, 2)
    elif isinstance(obj, dict):
        return {k: clean_json(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [clean_json(v) for v in obj]
    return obj

def parse_rank(value):
    if pd.isna(value):
        return np.nan
    if isinstance(value, (int, float)):
        return float(value)
    m = re.search(r"\d+", str(value))
    return float(m.group()) if m else np.nan

def build_data(excel_file=EXCEL_FILE):
    if not os.path.exists(excel_file):
        raise FileNotFoundError(f"{excel_file} not found.")

    print(f"Reading {excel_file}...")
    xls = pd.ExcelFile(excel_file)
    sheets = [s for s in xls.sheet_names if s.isdigit()]
    sheets.sort()
    
    all_data = {}
    master_df_list = []

    # 1. Load and Process Sheets
    for year_str in sheets:
        print(f"  Processing {year_str}...")
        df = pd.read_excel(xls, sheet_name=year_str)
        # Ensure score and metrics are numeric
        cols_to_fix = ['Score'] + [c for c in SUB_METRICS if c in df.columns]
        for col in cols_to_fix:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        if 'Rank' in df.columns:
            df['Rank'] = df['Rank'].apply(parse_rank)
        
        df = df.dropna(subset=['Score'])
        df['Year'] = int(year_str)
        all_data[year_str] = df
        master_df_list.append(df)

    master_df = pd.concat(master_df_list, ignore_index=True)
    years = [int(y) for y in sheets]
    latest_year = sheets[-1]
    df_latest = all_data[latest_year]

    # 2. Compute Analytics Object
    data_obj = {
        "years": years,
        "year_stats": [],
        "year_dist": {},
        "year_corr": {},
        "corr_matrix": {},
        "year_top_bot": {},
        "year_top10": {},
        "year_state_stats": {},
        "year_submetric_avgs": {},
        "year_scatter": {},
        "year_scatter_multi": {},
        "iit_traj": [],
        "iit_regression": {},
        "regression": {},
        "yoy_growth": [],
        "moving_avg_3": [],
        "feature_importance": {"rf": {}, "gb": {}, "lr_coef": {}, "lr_r2": 0},
        "metric_trend": {},
        "rank_mobility": {},
        "clustering": {"viz": [], "labels": {}, "profiles": {}, "pca_explained": []},
        "pca_variance": {"var_exp": [], "cumvar": []},
        "hypothesis": {"anova_state": {}, "ttest_top5": {}, "anova_year": {}},
        "outliers": {"z_method": [], "iqr_method": []},
        "rank_stability": [],
        "consistent_top": [],
        "state_dominance": {},
        "score_validation": [],
        "cluster_centers": [],
        "explorer": []
    }

    # Helper stats per year
    mean_scores = []
    for yr in sheets:
        df = all_data[yr]
        scores = df['Score']
        
        # stats
        stats = {
            "year": int(yr),
            "count": len(df),
            "state_count": df['State'].nunique(),
            "min": scores.min(),
            "max": scores.max(),
            "mean": scores.mean(),
            "median": scores.median(),
            "std": scores.std(),
            "variance": scores.var(),
            "skewness": skew(scores),
            "kurtosis": kurtosis(scores),
            "q1": scores.quantile(0.25),
            "q3": scores.quantile(0.75),
            "iqr": scores.quantile(0.75) - scores.quantile(0.25),
            "p10": scores.quantile(0.1),
            "p25": scores.quantile(0.25),
            "p50": scores.quantile(0.5),
            "p75": scores.quantile(0.75),
            "p90": scores.quantile(0.9)
        }
        data_obj["year_stats"].append(stats)
        mean_scores.append(scores.mean())

        # Distribution
        bins = [0, 40, 50, 55, 60, 65, 70, 75, 80, 85, 90, 100]
        labels = ["<40", "40-50", "50-55", "55-60", "60-65", "65-70", "70-75", "75-80", "80-85", "85-90", "90+"]
        dist = pd.cut(scores, bins=bins, labels=labels).value_counts().to_dict()
        data_obj["year_dist"][yr] = {k: int(v) for k, v in dist.items() if v > 0}

        # Correlation with Score
        available_metrics = [c for c in SUB_METRICS if c in df.columns]
        if available_metrics:
            corr_with_score = df[available_metrics].corrwith(df['Score']).sort_values(ascending=False).to_dict()
            data_obj["year_corr"][yr] = corr_with_score
            data_obj["corr_matrix"][yr] = df[available_metrics].corr().to_dict()

        # Top 25% vs Bottom 25%
        q25 = df['Score'].quantile(0.25)
        q75 = df['Score'].quantile(0.75)
        df_top = df[df['Score'] >= q75]
        df_bot = df[df['Score'] <= q25]
        tb_stats = {}
        for m in available_metrics:
            tb_stats[m] = {"top": df_top[m].mean(), "bot": df_bot[m].mean()}
        data_obj["year_top_bot"][yr] = tb_stats

        # Top 10
        t10 = df.nlargest(10, 'Score')[['Institute Name', 'City', 'State', 'Score', 'PR', 'Rank']]
        data_obj["year_top10"][yr] = t10.rename(columns={'Institute Name': 'name', 'City': 'city', 'State': 'state', 'Score': 'score', 'PR': 'pr', 'Rank': 'rank'}).to_dict(orient='records')

        # State Stats
        st_stats = df.groupby('State')['Score'].agg(['count', 'mean']).reset_index()
        st_stats = st_stats.sort_values('count', ascending=False)
        data_obj["year_state_stats"][yr] = st_stats.rename(columns={'State':'state', 'mean':'avg'}).to_dict(orient='records')

        # Sub-metric averages
        data_obj["year_submetric_avgs"][yr] = df[available_metrics].mean().to_dict()

        # Scatter (Default PR vs Score)
        if 'PR' in df.columns:
            pts = df[['PR', 'Score']].rename(columns={'PR':'x', 'Score':'y'}).to_dict(orient='records')
            data_obj["year_scatter"][yr] = {"points": pts, "xLabel": "PR", "corr": round(df['PR'].corr(df['Score']), 3)}

        # Multi Scatter
        multi = {}
        for m in available_metrics:
            pts = df[[m, 'Score']].rename(columns={m:'x', 'Score':'y'}).to_dict(orient='records')
            multi[m] = {"points": pts, "corr": round(df[m].corr(df['Score']), 3)}
        data_obj["year_scatter_multi"][yr] = multi

    available_metrics_latest = [c for c in SUB_METRICS if c in df_latest.columns]
    if not available_metrics_latest:
        available_metrics_latest = [c for c in ['TLR', 'RPC', 'GO', 'OI', 'PR', 'PCS'] if c in df_latest.columns]

    # 3. IIT Madras Track (Search by common ID across years)
    iit_id = "IR-E-U-0456"
    iit_records = master_df[master_df['Institute ID'] == iit_id].sort_values('Year')
    if not iit_records.empty:
        data_obj["iit_traj"] = iit_records[['Year', 'Score']].rename(columns={'Year':'year', 'Score':'score'}).to_dict(orient='records')
        # Forecast IIT
        X_iit = iit_records['Year'].values.reshape(-1, 1)
        y_iit = iit_records['Score'].values
        if len(X_iit) > 1:
            lr_iit = LinearRegression().fit(X_iit, y_iit)
            data_obj["iit_regression"] = {
                "slope": round(float(lr_iit.coef_[0]), 3),
                "pred_2026": round(float(lr_iit.predict([[2026]])[0]), 2),
                "pred_2027": round(float(lr_iit.predict([[2027]])[0]), 2)
            }

    # 4. Global Regression & Forecast
    X_yr = np.array(years).reshape(-1, 1)
    y_yr = np.array(mean_scores)
    lr_yr = LinearRegression().fit(X_yr, y_yr)
    data_obj["regression"] = {
        "slope": round(float(lr_yr.coef_[0]), 3),
        "r2": round(lr_yr.score(X_yr, y_yr), 3),
        "pred_2026": round(float(lr_yr.predict([[2026]])[0]), 2),
        "pred_2027": round(float(lr_yr.predict([[2027]])[0]), 2)
    }

    # 5. YoY Growth & Moving Avg
    yoy = []
    for i in range(1, len(mean_scores)):
        growth = ((mean_scores[i] - mean_scores[i-1]) / mean_scores[i-1]) * 100
        yoy.append({"period": f"{years[i-1]}-{years[i]}", "growth_pct": round(growth, 2)})
    data_obj["yoy_growth"] = yoy
    
    mavg = []
    for i in range(len(mean_scores)):
        if i < 2: mavg.append(None)
        else: mavg.append(round(np.mean(mean_scores[i-2:i+1]), 2))
    data_obj["moving_avg_3"] = mavg

    # 6. ML Feature Importance (Latest Year)
    X = df_latest[available_metrics_latest].fillna(0)
    y = df_latest['Score']
    rf = RandomForestRegressor(n_estimators=50, n_jobs=1, random_state=42).fit(X, y)
    gb = GradientBoostingRegressor(n_estimators=50, random_state=42).fit(X, y)
    lr = LinearRegression().fit(X, y)
    
    data_obj["feature_importance"]["rf"] = dict(zip(available_metrics_latest, rf.feature_importances_))
    data_obj["feature_importance"]["gb"] = dict(zip(available_metrics_latest, gb.feature_importances_))
    data_obj["feature_importance"]["lr_coef"] = dict(zip(available_metrics_latest, lr.coef_))
    data_obj["feature_importance"]["lr_r2"] = float(lr.score(X, y))

    # Metric trend evolution across years
    for m in available_metrics_latest:
        series = []
        for yr in years:
            ydf = all_data[str(yr)]
            if m in ydf.columns:
                series.append({"year": int(yr), "avg": float(ydf[m].mean())})
        data_obj["metric_trend"][m] = series

    # 7. Clustering & PCA (Latest Year)
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    
    kmeans = KMeans(n_clusters=3, random_state=42, n_init=10).fit(X_scaled)
    df_latest['cluster'] = kmeans.labels_
    
    pca_full = PCA().fit(X_scaled)
    var_exp = (pca_full.explained_variance_ratio_ * 100).tolist()
    data_obj["pca_variance"]["var_exp"] = var_exp
    data_obj["pca_variance"]["cumvar"] = np.cumsum(var_exp).tolist()

    pca = PCA(n_components=2)
    pca_res = pca.fit_transform(X_scaled)
    df_latest['pca1'] = pca_res[:, 0]
    df_latest['pca2'] = pca_res[:, 1]
    data_obj["clustering"]["pca_explained"] = (pca.explained_variance_ratio_ * 100).round(2).tolist()
    
    # Cluster Profiles
    for i in range(3):
        c_df = df_latest[df_latest['cluster'] == i]
        data_obj["clustering"]["profiles"][i] = {
            "count": int(len(c_df)),
            "mean_score": float(c_df['Score'].mean()) if len(c_df) else 0
        }
        data_obj["cluster_centers"].append({
            "id": i,
            "name": f"Cluster {i+1}",
            "count": len(c_df),
            "score": c_df['Score'].mean(),
            "metrics": c_df[available_metrics_latest].mean().to_dict()
        })
    data_obj["clustering"]["labels"] = {0: "High Performers", 1: "Balanced Performers", 2: "Emerging Performers"}
    data_obj["clustering"]["viz"] = df_latest[['pca1', 'pca2', 'Score', 'cluster']].rename(
        columns={'pca1': 'x', 'pca2': 'y', 'Score': 'score'}
    ).to_dict(orient='records')

    # 8. Rank mobility
    for i in range(1, len(years)):
        y_prev = str(years[i - 1])
        y_curr = str(years[i])
        left = all_data[y_prev][['Institute ID', 'Institute Name', 'Rank']].rename(columns={'Rank': 'rank_prev'})
        right = all_data[y_curr][['Institute ID', 'Institute Name', 'Rank']].rename(columns={'Rank': 'rank_curr'})
        merged = pd.merge(left, right, on='Institute ID', how='inner', suffixes=('_prev', '_curr'))
        merged['delta'] = merged['rank_prev'] - merged['rank_curr']
        merged = merged.dropna(subset=['delta'])
        if merged.empty:
            data_obj["rank_mobility"][f"{y_prev}-{y_curr}"] = {"improvers": [], "decliners": [], "all_deltas": []}
            continue
        improvers = merged.sort_values('delta', ascending=False).head(10)
        decliners = merged.sort_values('delta', ascending=True).head(10)
        data_obj["rank_mobility"][f"{y_prev}-{y_curr}"] = {
            "improvers": improvers[['Institute Name_prev', 'delta']].rename(columns={'Institute Name_prev': 'name'}).assign(
                delta=lambda d: d['delta'].round(0).astype(int)
            ).to_dict(orient='records'),
            "decliners": decliners[['Institute Name_prev', 'delta']].rename(columns={'Institute Name_prev': 'name'}).assign(
                delta=lambda d: d['delta'].round(0).astype(int)
            ).to_dict(orient='records'),
            "all_deltas": merged['delta'].round(0).astype(int).tolist()
        }

    # 9. Hypothesis tests
    latest_scores = all_data[latest_year]
    state_groups = [g['Score'].values for _, g in latest_scores.groupby('State') if len(g) >= 2]
    if len(state_groups) >= 2:
        f_stat, p_val = f_oneway(*state_groups)
        data_obj["hypothesis"]["anova_state"] = {"f": float(f_stat), "p": float(p_val), "significant": bool(p_val < 0.05)}

    state_avg = latest_scores.groupby('State')['Score'].mean().sort_values(ascending=False)
    top5_states = state_avg.head(5).index.tolist()
    grp_top5 = latest_scores[latest_scores['State'].isin(top5_states)]['Score'].values
    grp_rest = latest_scores[~latest_scores['State'].isin(top5_states)]['Score'].values
    if len(grp_top5) > 2 and len(grp_rest) > 2:
        t_stat, p_val = ttest_ind(grp_top5, grp_rest, equal_var=False)
        data_obj["hypothesis"]["ttest_top5"] = {
            "t": float(t_stat), "p": float(p_val), "significant": bool(p_val < 0.05),
            "mean_top5": float(np.mean(grp_top5)), "mean_rest": float(np.mean(grp_rest))
        }

    year_groups = [all_data[str(y)]['Score'].values for y in years if len(all_data[str(y)]) >= 2]
    if len(year_groups) >= 2:
        f_stat, p_val = f_oneway(*year_groups)
        data_obj["hypothesis"]["anova_year"] = {"f": float(f_stat), "p": float(p_val), "significant": bool(p_val < 0.05)}

    # 10. Outliers (latest year)
    scores_latest = latest_scores['Score']
    z = (scores_latest - scores_latest.mean()) / scores_latest.std(ddof=0)
    z_out = latest_scores[np.abs(z) > 2][['Institute Name', 'Score']].rename(columns={'Institute Name': 'name', 'Score': 'score'})
    q1, q3 = scores_latest.quantile(0.25), scores_latest.quantile(0.75)
    iqr = q3 - q1
    upper = q3 + 1.5 * iqr
    iqr_out = latest_scores[latest_scores['Score'] > upper][['Institute Name', 'Score']].rename(columns={'Institute Name': 'name', 'Score': 'score'})
    data_obj["outliers"]["z_method"] = z_out.to_dict(orient='records')
    data_obj["outliers"]["iqr_method"] = iqr_out.to_dict(orient='records')

    # 11. Rank stability + consistent top
    rank_df = master_df[['Institute ID', 'Institute Name', 'Rank']].copy()
    rank_df = rank_df.dropna(subset=['Rank'])
    stab = rank_df.groupby(['Institute ID', 'Institute Name'])['Rank'].agg(['mean', 'std', 'count']).reset_index()
    stab = stab[stab['count'] >= 3].sort_values('std')
    data_obj["rank_stability"] = stab.rename(columns={
        'Institute Name': 'name', 'mean': 'mean_rank', 'std': 'std_rank'
    })[['name', 'mean_rank', 'std_rank']].to_dict(orient='records')

    top_counts = master_df[master_df['Rank'] <= 10].groupby('Institute Name').size().sort_values(ascending=False).head(15)
    data_obj["consistent_top"] = [[name, int(cnt)] for name, cnt in top_counts.items()]

    # 12. State dominance by year
    for y in years:
        ydf = all_data[str(y)]
        st = ydf.groupby('State')['Score'].agg(['count', 'mean']).reset_index().rename(columns={'State': 'state', 'mean': 'avg'})
        st['dominance'] = (st['count'] * st['avg']) / 100.0
        st = st.sort_values('dominance', ascending=False).head(12)
        data_obj["state_dominance"][str(y)] = st.to_dict(orient='records')

    # 13. Score validation table (latest year top ranks)
    val = latest_scores.sort_values('Score', ascending=False).head(20)
    data_obj["score_validation"] = val[['Rank', 'Institute Name', 'Score']].rename(
        columns={'Rank': 'rank', 'Institute Name': 'name', 'Score': 'actual'}
    ).to_dict(orient='records')

    # 14. Explorer (Full)
    exp = master_df[['Rank', 'Institute ID', 'Institute Name', 'City', 'State', 'Score', 'PR', 'Year']]
    data_obj["explorer"] = exp.rename(columns={
        'Rank':'rank', 'Institute ID':'id', 'Institute Name':'name', 'City':'city', 'State':'state', 'Score':'score', 'PR':'pr', 'Year':'year'
    }).to_dict(orient='records')

    # Final Clean
    data_obj = clean_json(data_obj)
    return data_obj

def sync():
    try:
        data_obj = build_data(EXCEL_FILE)
    except FileNotFoundError as exc:
        print(f"Error: {exc}")
        return
    
    json_str = json.dumps(data_obj, indent=2)
    
    print(f"Injecting data into {HTML_FILE}...")
    with open(HTML_FILE, 'r', encoding='utf-8') as f:
        content = f.read()
    
    new_content = re.sub(r'const DATA\s*=\s*\{.*?\};', f'const DATA = {json_str};', content, flags=re.DOTALL)
    
    with open(HTML_FILE, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print("Dashboard synchronized successfully!")

if __name__ == "__main__":
    sync()
