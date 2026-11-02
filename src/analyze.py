"""Case study analysis: mining patient records for common patterns.

Runs the full pipeline used in the case study report:
  1. Load UCI Cleveland heart-disease records (303 patients).
  2. Clean: impute missing values, encode nominal attributes, binarise target.
  3. Cluster (K-means) on 5 continuous clinical measures -> common patient groups.
  4. Classify (decision tree) -> predict heart disease, evaluate on a hold-out set.

Outputs: reports/results.json, reports/figures/*.png
Re-run with:  python src/analyze.py
"""

from __future__ import annotations

import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.metrics import (
    accuracy_score,
    confusion_matrix,
    f1_score,
    precision_score,
    recall_score,
    silhouette_score,
)
from sklearn.model_selection import cross_val_score, train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.tree import DecisionTreeClassifier, export_text, plot_tree

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw" / "processed.cleveland.data"
OUT = ROOT / "reports"
FIG = OUT / "figures"

COLUMNS = [
    "age", "sex", "cp", "trestbps", "chol", "fbs", "restecg",
    "thalach", "exang", "oldpeak", "slope", "ca", "thal", "num",
]

CP_LABELS = {
    1: "typical angina",
    2: "atypical angina",
    3: "non-anginal",
    4: "asymptomatic",
}

# Continuous features used for clustering (the five plain clinical measures).
CLUSTER_FEATURES = ["age", "trestbps", "chol", "thalach", "oldpeak"]

plt.rcParams.update({
    "font.size": 10.5,
    "axes.titlesize": 12,
    "axes.titleweight": "bold",
    "axes.grid": True,
    "grid.alpha": 0.3,
    "figure.dpi": 120,
})


def load_and_clean() -> tuple[pd.DataFrame, dict]:
    df = pd.read_csv(RAW, header=None, names=COLUMNS, na_values="?")
    n_records = len(df)

    missing = {c: int(df[c].isna().sum()) for c in COLUMNS if df[c].isna().any()}
    imputed = {}
    for col in missing:
        median = float(df[col].median())
        df[col] = df[col].fillna(median)
        imputed[col] = median

    # Binary target: 0 = no disease, 1-4 = disease present.
    df["target"] = (df["num"] > 0).astype(int)

    cleaning = {
        "records_before": n_records,
        "records_after": int(len(df)),
        "missing_values": missing,
        "median_imputed": imputed,
        "duplicate_rows": int(df.duplicated().sum()),
        "class_counts": {
            "no_disease": int((df["target"] == 0).sum()),
            "disease": int((df["target"] == 1).sum()),
        },
    }
    return df, cleaning


def clustering(df: pd.DataFrame, fig_dir: Path) -> dict:
    X = df[CLUSTER_FEATURES]
    scaler = StandardScaler()
    Xs = scaler.fit_transform(X)

    inertias, silhouettes = {}, {}
    for k in range(1, 9):
        km = KMeans(n_clusters=k, n_init=10, random_state=42)
        km.fit(Xs)
        inertias[k] = float(km.inertia_)
        if k >= 2:
            silhouettes[k] = float(silhouette_score(Xs, km.labels_))

    best_k = max(silhouettes, key=silhouettes.get)
    kmeans = KMeans(n_clusters=best_k, n_init=10, random_state=42)
    labels = kmeans.fit_predict(Xs)
    overall_sil = float(silhouette_score(Xs, labels))

    # Cluster profiles in original units.
    profile_src = df[CLUSTER_FEATURES].copy()
    profile_src["cluster"] = labels
    centers = profile_src.groupby("cluster").agg(["mean", "size"])
    profiles = []
    for c in range(best_k):
        profiles.append({
            "cluster": int(c),
            "size": int(centers.loc[c, ("age", "size")]),
            **{f: round(float(centers.loc[c, (f, "mean")]), 2)
               for f in CLUSTER_FEATURES},
        })

    # Figure 1: elbow + silhouette
    fig, axes = plt.subplots(1, 2, figsize=(8.0, 3.2))
    ks = list(map(int, inertias))
    axes[0].plot(ks, [inertias[k] for k in ks], "o-", color="#1f3864")
    axes[0].axvline(best_k, color="#c00000", linestyle="--", linewidth=1)
    axes[0].set_xlabel("Number of clusters (k)", fontsize=11.5)
    axes[0].set_ylabel("Inertia (within-cluster sum of squares)", fontsize=11.5)
    axes[0].set_title("Elbow curve")
    sk = list(map(int, silhouettes))
    axes[1].bar(sk, [silhouettes[k] for k in sk], color="#2e75b6")
    axes[1].axvline(best_k, color="#c00000", linestyle="--", linewidth=1)
    axes[1].set_xlabel("Number of clusters (k)", fontsize=11.5)
    axes[1].set_ylabel("Silhouette score", fontsize=11.5)
    axes[1].set_title(f"Silhouette score (best k = {best_k})")
    axes[1].set_ylim(0, max(silhouettes.values()) * 1.25)
    for k, v in zip(sk, [silhouettes[k] for k in sk]):
        axes[1].text(k, v, f"{v:.2f}", ha="center", va="bottom", fontsize=9.5)
    fig.tight_layout()
    fig.savefig(fig_dir / "fig1_elbow_silhouette.png", bbox_inches="tight", dpi=300)
    plt.close(fig)

    # Figure 2: PCA scatter of clusters
    pca = PCA(n_components=2, random_state=42)
    Z = pca.fit_transform(Xs)
    colors = ["#1f3864", "#c00000", "#548235", "#7030a0", "#bf8f00"]
    fig, ax = plt.subplots(figsize=(6.4, 4.4))
    for c in range(best_k):
        m = labels == c
        ax.scatter(Z[m, 0], Z[m, 1], s=34, alpha=0.75,
                   color=colors[c % len(colors)],
                   label=f"Cluster {c + 1} (n={int(m.sum())})")
    ax.set_xlabel(f"Principal component 1 ({pca.explained_variance_ratio_[0]*100:.1f}% variance)")
    ax.set_ylabel(f"Principal component 2 ({pca.explained_variance_ratio_[1]*100:.1f}% variance)")
    ax.set_title("Patient groups found by K-means")
    ax.legend(frameon=True, fontsize=9)
    fig.tight_layout()
    fig.savefig(fig_dir / "fig2_clusters.png", bbox_inches="tight", dpi=300)
    plt.close(fig)

    pca_explained = [round(float(v) * 100, 1) for v in pca.explained_variance_ratio_]

    # Disease rate per cluster: links the unsupervised groups to real outcomes.
    ct = pd.crosstab(labels, df["target"], normalize="index")
    disease_rate = {
        int(idx): (round(float(row[1]) * 100, 1) if 1 in row.index else 0.0)
        for idx, row in ct.iterrows()
    }

    # Worked example: distance of the first patient to each centroid (scaled space).
    patient = Xs[0]
    dists = [float(np.linalg.norm(patient - kmeans.cluster_centers_[i]))
             for i in range(best_k)]
    assigned = int(np.argmin(dists))

    return {
        "features": CLUSTER_FEATURES,
        "scaler_mean": {f: round(float(scaler.mean_[i]), 2)
                        for i, f in enumerate(CLUSTER_FEATURES)},
        "scaler_scale": {f: round(float(scaler.scale_[i]), 2)
                         for i, f in enumerate(CLUSTER_FEATURES)},
        "k_inertia": {int(k): round(v, 1) for k, v in inertias.items()},
        "k_silhouette": {int(k): round(v, 4) for k, v in silhouettes.items()},
        "chosen_k": int(best_k),
        "silhouette": round(overall_sil, 4),
        "pca_explained_variance_pct": pca_explained,
        "cluster_sizes": [p["size"] for p in profiles],
        "disease_rate_pct": disease_rate,
        "profiles": profiles,
        "centroids_scaled": [[round(float(v), 3) for v in row]
                             for row in kmeans.cluster_centers_],
        "centroids_unscaled": [[round(float(v), 2) for v in row]
                               for row in scaler.inverse_transform(kmeans.cluster_centers_)],
        "worked_example": {
            "patient": {f: float(df[CLUSTER_FEATURES].iloc[0][f])
                        for f in CLUSTER_FEATURES},
            "scaled": [round(float(v), 2) for v in patient],
            "distances": [round(d, 3) for d in dists],
            "assigned_cluster": assigned,
        },
    }


def classification(df: pd.DataFrame, fig_dir: Path) -> dict:
    # One-hot encode nominal attributes so tree splits mean something.
    nominal = ["cp", "restecg", "slope", "ca", "thal"]
    features = df.drop(columns=["num", "target"])
    X = pd.get_dummies(features, columns=nominal, dtype=int)
    y = df["target"]
    feature_names = list(X.columns)

    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.30, stratify=y, random_state=42
    )

    tree = DecisionTreeClassifier(
        max_depth=4, min_samples_leaf=8, random_state=42
    )
    tree.fit(X_train, y_train)
    y_pred = tree.predict(X_test)

    metrics = {
        "accuracy": round(float(accuracy_score(y_test, y_pred)), 4),
        "precision": round(float(precision_score(y_test, y_pred)), 4),
        "recall": round(float(recall_score(y_test, y_pred)), 4),
        "f1": round(float(f1_score(y_test, y_pred)), 4),
    }
    tn, fp, fn, tp = (int(v) for v in confusion_matrix(y_test, y_pred).ravel())
    baseline = round(float(max(y_test.mean(), 1 - y_test.mean())), 4)

    cv = cross_val_score(
        DecisionTreeClassifier(max_depth=4, min_samples_leaf=8, random_state=42),
        X, y, cv=10,
    )
    metrics["cv10_mean"] = round(float(cv.mean()), 4)
    metrics["cv10_std"] = round(float(cv.std()), 4)
    metrics["baseline_majority"] = baseline
    metrics["confusion"] = [[tn, fp], [fn, tp]]
    metrics["train_size"] = int(len(X_train))
    metrics["test_size"] = int(len(X_test))
    metrics["n_features_after_encoding"] = int(X.shape[1])

    importances = sorted(
        zip(feature_names, tree.feature_importances_),
        key=lambda kv: kv[1], reverse=True,
    )
    top_features = [{"feature": f, "importance": round(float(v), 3)}
                    for f, v in importances[:6] if v > 0]

    rules_text = []
    for line in export_text(tree, feature_names=feature_names, decimals=0).splitlines()[:40]:
        if line.strip():
            rules_text.append(line)

    # Worked example: one test-set patient (never seen during training) and the
    # exact path the tree walks to classify them.
    example_idx = None
    for pos, idx in enumerate(X_test.index):
        if y_pred[pos] == y_test.iloc[pos] == 1:
            example_idx = pos
            break
    if example_idx is None:
        example_idx = 0
    row_pos = X_test.index[example_idx]
    patient_row = X_test.iloc[example_idx]
    path_steps = []
    node = 0
    tree_struct = tree.tree_
    while tree_struct.children_left[node] != tree_struct.children_right[node]:
        f = tree_struct.feature[node]
        t = float(tree_struct.threshold[node])
        go_left = patient_row.iloc[f] <= t
        path_steps.append({
            "feature": feature_names[f],
            "threshold": round(t, 2),
            "value": round(float(patient_row.iloc[f]), 2),
            "go": "left" if go_left else "right",
        })
        node = (tree_struct.children_left[node] if go_left
                else tree_struct.children_right[node])
    raw = df.loc[row_pos]
    tree_example = {
        "record_index": int(row_pos),
        "raw": {c: (None if pd.isna(raw[c]) else float(raw[c]))
                for c in df.columns if c not in ("num", "target")},
        "path": path_steps,
        "actual": int(y_test.iloc[example_idx]),
        "predicted": int(y_pred[example_idx]),
    }

    # Figure 3: confusion matrix
    fig, ax = plt.subplots(figsize=(4.6, 4.0))
    cm = np.array([[tn, fp], [fn, tp]])
    im = ax.imshow(cm, cmap="Blues")
    ax.set_xticks([0, 1], ["No disease", "Disease"])
    ax.set_yticks([0, 1], ["No disease", "Disease"])
    ax.set_xlabel("Predicted")
    ax.set_ylabel("Actual")
    ax.set_title("Confusion matrix (test set)")
    for (i, j), val in np.ndenumerate(cm):
        ax.text(j, i, str(val), ha="center", va="center",
                color="white" if val > cm.max() / 2 else "black",
                fontsize=15, fontweight="bold")
    fig.colorbar(im, fraction=0.046, pad=0.04)
    fig.tight_layout()
    fig.savefig(fig_dir / "fig3_confusion_matrix.png", bbox_inches="tight", dpi=300)
    plt.close(fig)

    # Figure 4: the decision tree itself (shown on a landscape page in the
    # report). Node font is deliberately small: 13 leaves must fit across one
    # printed landscape width without their text overlapping.
    fig, ax = plt.subplots(figsize=(11, 6.6))
    plot_tree(
        tree, feature_names=feature_names, class_names=["no", "yes"],
        filled=True, rounded=True, fontsize=7.5, ax=ax,
        proportion=False, impurity=True,
    )
    ax.set_title("Decision tree learned from patient records (depth <= 4)",
                 fontsize=11)
    fig.tight_layout()
    fig.savefig(fig_dir / "fig4_decision_tree.png", bbox_inches="tight", dpi=300)
    plt.close(fig)

    return {
        "metrics": metrics,
        "top_features": top_features,
        "tree_rules": rules_text,
        "tree_depth": int(tree.get_depth()),
        "tree_leaves": int(tree.get_n_leaves()),
        "tree_example": tree_example,
    }


def sample_rows(df: pd.DataFrame, n: int = 10) -> list[dict]:
    rows = []
    for _, r in df.head(n).iterrows():
        rows.append({
            "age": int(r["age"]),
            "sex": "M" if r["sex"] == 1 else "F",
            "cp": CP_LABELS.get(int(r["cp"]), str(int(r["cp"]))),
            "trestbps": int(r["trestbps"]),
            "chol": int(r["chol"]),
            "thalach": int(r["thalach"]),
            "disease": "Yes" if r["target"] == 1 else "No",
        })
    return rows


def descriptive_stats(df: pd.DataFrame) -> list[dict]:
    rows = []
    for col in CLUSTER_FEATURES:
        s = df[col]
        rows.append({
            "attribute": col,
            "mean": round(float(s.mean()), 1),
            "std": round(float(s.std()), 1),
            "min": round(float(s.min()), 1),
            "max": round(float(s.max()), 1),
        })
    return rows


def main() -> None:
    FIG.mkdir(parents=True, exist_ok=True)
    df, cleaning = load_and_clean()
    print(f"Loaded {len(df)} patient records")
    print(f"Missing values found: {cleaning['missing_values']}")

    clus = clustering(df, FIG)
    print(f"Chosen k: {clus['chosen_k']}  silhouette={clus['silhouette']}")
    cls = classification(df, FIG)
    print(f"Tree metrics: {cls['metrics']}")

    results = {
        "dataset": {
            "name": "UCI Heart Disease (Cleveland processed)",
            "records": int(len(df)),
            "raw_attributes": 13,
            "target": "num > 0 => heart disease present",
            "source": "https://archive.ics.uci.edu/dataset/45/heart+disease",
        },
        "cleaning": cleaning,
        "descriptives": descriptive_stats(df),
        "sample_rows": sample_rows(df),
        "clustering": clus,
        "classification": cls,
    }
    out = OUT / "results.json"
    out.write_text(json.dumps(results, indent=2))
    print(f"Wrote {out.relative_to(ROOT)} and {len(list(FIG.glob('*.png')))} figures")


if __name__ == "__main__":
    main()
