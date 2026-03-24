import os
import time

import numpy as np
import pandas as pd
import xlwings as xw
from sklearn.dummy import DummyRegressor
from sklearn.gaussian_process import GaussianProcessRegressor
from sklearn.gaussian_process.kernels import WhiteKernel, ConstantKernel, DotProduct, RBF
from sklearn.inspection import permutation_importance
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_absolute_error, r2_score, mean_squared_error

from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import ElasticNetCV, RidgeCV

from sklearn.experimental import enable_halving_search_cv
from sklearn.model_selection import HalvingRandomSearchCV, RandomizedSearchCV

def create_clean_df(df, mask):
    # Filterung:
    r = df["TRIMMEN_START"]
    mask_leer = r.isna()

    dt = pd.to_datetime(r, format="%d/%m/%Y %H:%M", dayfirst=True, errors="coerce")
    grenze = pd.Timestamp('2023-01-01 00:00:00')

    mask_ab_2023 = dt >= grenze

    mask_ab_2023 = mask_ab_2023 | mask_leer

    # Lasernummern vereinfachen
    laser_map = {12: 3, 17: 4, 18: 5, 20: 6, 121: 7}
    for col in ["LASER_T", "LASER_NT", "LASER_NM"]:
        if col in df.columns:
            s = pd.to_numeric(df[col], errors="coerce")
            df[col] = s.replace(laser_map)

    mask = mask & mask_ab_2023 & (df["NC_CALC"] != "N")

    df_encoded = df.copy()

    # Kategorische Spalten
    categorical_cols = ["TYPTC", "TYPTCA", "LADP", "TOLP", "NT_NM"]

    # One-Hot-Encoding
    df_encoded = pd.get_dummies(
        df_encoded,
        columns=categorical_cols,
        prefix=categorical_cols,
        prefix_sep="_",
        drop_first=False,  # alle Kategorien behalten
        dtype=int
    )

    # Ergebnis
    print(df_encoded.head())
    print(df_encoded.shape)

    return df_encoded.loc[mask].copy()

def select_features():
    print("Starte Random Forest Regressor...")
    excel_path = "C:\\Users\\pythonProject\\NC_CALC_Programm.xlsm"

    # Funktionsaufruf
    try:
        wb = xw.Book.caller()
    except Exception:
        wb = xw.Book(excel_path)

    ws_dateiaus = wb.sheets["DateiAus"]  # Ausgabe-Sheet

    ws_dateiaus.cells(2, 11).value = "Random Forest Regressor"

    df = pd.read_excel(wb.fullname, sheet_name="Rohdaten")
    print("Daten gefunden. Beginne Modelltraining...")

    mask = (
            df["NC_T"].notna() &
            df["FT_T"].notna()
    )

    df_clean = create_clean_df(df, mask)

    drop_cols = [
        "TRIMMEN_START",
        "NC_CALC",
        "NC_CALC_DATE",
        "eac",
        "FA",
        "L1_TDS", "L2_TDS", "L3_TDS",
        "L4_TDS", "L5_TDS", "L6_TDS", "L7_TDS",
        "BEMERKUNG"
    ]

    target = "NC_T"

    df_model = df_clean.copy()

    # Zielvariable
    y = df_model[target]

    # Spalten entfernen
    X = df_model.drop(columns=drop_cols + [target], errors="ignore")

    # Nur numerische Features
    X = X.select_dtypes(include=[np.number])

    print("Anzahl Features:", X.shape[1])

    # Pearson-Korrelation mit RFR:
    X_train, X_test, y_train, y_test = train_test_split(
        X, y,
        test_size=0.2,
        random_state=42
    )

    rf = RandomForestRegressor(
        n_estimators=500,
        random_state=42,
        n_jobs=-1
    )

    rf.fit(X_train, y_train)

    y_pred = rf.predict(X_test)

    print("R²:", r2_score(y_test, y_pred))
    print("MAE:", mean_absolute_error(y_test, y_pred))

    corr_pearson = pd.Series(
        {col: X[col].corr(y) for col in X.columns},
        name="Corr_Pearson"
    ).sort_values(ascending=False)

    print(corr_pearson.head(20))

    # Feature Importance
    mdi_importance = pd.DataFrame({
        "Feature": X.columns,
        "MDI_Importance": rf.feature_importances_
    }).sort_values(by="MDI_Importance", ascending=False)

    print(mdi_importance.head(20))

    perm_importance = permutation_importance(
        rf,
        X_test,
        y_test,
        n_repeats=10,
        random_state=42,
        n_jobs=-1
    )

    perm_df = pd.DataFrame({
        "Feature": X.columns,
        "Permutation_Importance_Mean": perm_importance.importances_mean,
        "Permutation_Importance_STD": perm_importance.importances_std
    }).sort_values(by="Permutation_Importance_Mean", ascending=False)

    with pd.option_context("display.max_rows", None,
                           "display.max_columns", None):
        print(
            perm_df.sort_values(
                by="Permutation_Importance_Mean",
                ascending=False
            )
        )


def run_model():
    print("Starte Random Forest Regressor...")
    excel_path = "C:\\Users\\pythonProject\\NC_CALC_Programm.xlsm"

    try:
        wb = xw.Book.caller()
    except Exception:
        wb = xw.Book(excel_path)

    ws_auftraege = wb.sheets["Auftragstabelle"]     # Auftragstabelle
    ws_dateiaus = wb.sheets["DateiAus"]             # Ausgabe-Sheet

    ws_dateiaus.cells(2, 11).value = "Random Forest Regressor"

    df = pd.read_excel(wb.fullname, sheet_name="Rohdaten")
    print("Daten gefunden. Beginne Modelltraining...")

    mask_t = (
            df["NC_T"].notna() &
            df["FT_T"].notna()
    )

    mask_nt = (
            df["NC_NT"].notna() &
            df["FT_NT"].notna()
    )

    df_clean_t = create_clean_df(df, mask_t)
    df_clean_nt = create_clean_df(df, mask_nt)

    # Feature-Auswahl
    feature_cols = ["X_TEIL_ID", "ZW_NC", "R0_SEL_K"]
    feature_cols_t = ["FT_T", "NT_NM_NT/NM", "ZIELWERT_T"]
    feature_cols_nt = ["FT_NT", "TYPTCA_1.2003.1EG", "ZIELWERT_NT"]

    feature_cols_t = feature_cols_t + feature_cols
    feature_cols_nt = feature_cols_nt + feature_cols

    def train_rf(
            df_model: pd.DataFrame,
            train_feature_cols: list[str],
            target_col: str,
            label: str,
            *,
            random_state: int = 42
    ) -> RandomForestRegressor:
        X = df_model[train_feature_cols]
        y = df_model[target_col]

        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=0.2, random_state=random_state
        )

        base_rf = RandomForestRegressor(
            random_state=random_state,
            n_jobs=-1,
            bootstrap=True,
        )

        param_dist_halving_trim = {
            "n_estimators": [300, 450, 600],
            "max_depth": [4, 6, 8, 10, 12],
            "min_samples_split": [5, 10, 20, 30],
            "min_samples_leaf": [2, 4, 6, 8, 10],
            "max_features": [0.5, 0.7, 0.9, 1.0],
        }

        param_dist_halving_nt = {
            "n_estimators": [300, 450, 600],
            "max_depth": [3, 4, 6, 8, 10],
            "min_samples_split": [5, 10, 20, 30],
            "min_samples_leaf": [4, 6, 8, 12, 16, 20],
            "max_features": [0.5, 0.7, 0.9, 1.0],
        }

        if target_col == "NC_T":
            param_dist_halving = param_dist_halving_trim
            cv_halving = 3
            factor_halving = 3
        else:
            param_dist_halving = param_dist_halving_nt
            cv_halving = 5
            factor_halving = 2

        # RandomizedSearch darf bootstrap suchen
        param_dist_random = {
            **param_dist_halving,
            "bootstrap": [True, False],
        }

        # HalvingRandomSearchCV mit RandomSearchCV als Fallback
        try:
            halving = HalvingRandomSearchCV(
                estimator=base_rf,
                param_distributions=param_dist_halving,
                factor=factor_halving,
                resource="max_samples",
                max_resources=400,
                min_resources=150,
                scoring="neg_mean_absolute_error",
                cv=cv_halving,
                random_state=random_state,
                n_jobs=-1,
                verbose=0,
                refit=True
            )
            halving.fit(X_train, y_train)
            best_est = halving.best_estimator_
            best_params = halving.best_params_
            print(f"[{label}] HalvingRandomSearch best params: {best_params}")

        except Exception as e:
            print(
                f"[{label}] HalvingRandomSearch fehlgeschlagen ({type(e).__name__}: {e}) -> Fallback RandomizedSearch")

            rnd = RandomizedSearchCV(
                estimator=base_rf,
                param_distributions=param_dist_random,
                n_iter=25,
                scoring="neg_mean_absolute_error",
                cv=3,
                random_state=random_state,
                n_jobs=-1,
                verbose=0,
                refit=True
            )
            rnd.fit(X_train, y_train)
            best_est = rnd.best_estimator_
            best_params = rnd.best_params_
            print(f"[{label}] RandomizedSearch best params: {best_params}")

        # Bewertung der Modelle
        y_pred = best_est.predict(X_test)
        mae = mean_absolute_error(y_test, y_pred)
        rmse = np.sqrt(mean_squared_error(y_test, y_pred))
        r2 = r2_score(y_test, y_pred)

        print(f"[{label}] Zeilen: {len(df_model)}")
        print(f"[{label}] MAE: {mae:.4f} | RMSE: {rmse:.4f} | R²: {r2:.4f}")

        return best_est

    df_t = df_clean_t[["NC_T"] + feature_cols_t]
    df_nt = df_clean_nt[["NC_NT"] + feature_cols_nt]

    # Modelltraining Trimmen und Nachtrimmen
    rf_t = train_rf(df_t, feature_cols_t, "NC_T", "RF_Trimmen")
    rf_nt = train_rf(df_nt, feature_cols_nt, "NC_NT", "RF_Nachtrimmen")

    # Importance der Features
    importances = rf_t.feature_importances_
    feature_imp_df = pd.DataFrame({'Feature': feature_cols_t, 'Gini Importance': importances}).sort_values(
        'Gini Importance', ascending=False)
    print(feature_imp_df)

    importances = rf_nt.feature_importances_
    feature_imp_df = pd.DataFrame({'Feature': feature_cols_nt, 'Gini Importance': importances}).sort_values(
        'Gini Importance', ascending=False)
    print(feature_imp_df)

    # Auftragstabelle
    auftraege = []
    row = 2  # Zeile 1 = Überschriften

    while True:
        eac = ws_auftraege.cells(row, 2).value
        if eac is None:
            break
        x_teil_id = ws_auftraege.cells(row, 4).value
        nennwert = ws_auftraege.cells(row, 6).value
        NTNM = ws_auftraege.cells(row, 8).value

        auftraege.append({
            "row": row,
            "EAC": eac,
            "X_TEIL_ID": x_teil_id,
            "Nennwert": nennwert,
            "NTNM": NTNM,
        })
        row += 1

    anzahl_auftraege = len(auftraege)
    print(f"Anzahl Aufträge gefunden: {anzahl_auftraege}")
    print("Beginne Berechnung der Vorhersagen...")

    # Iteration durch jeden Auftrag in Auftragstabelle
    for idx, auftrag in enumerate(auftraege):
        eac = auftrag["EAC"]
        x_teil_id = auftrag["X_TEIL_ID"]
        nennwert = auftrag["Nennwert"]
        NTNM = auftrag["NTNM"]
        ntnm_Wert = 0

        if NTNM == "NTNM":
            ntnm_Wert = 1

        row = df.loc[df["EAC"].astype(str) == str(int(eac))]

        tn = 0
        typtca = ""
        if not row.empty:
            typtca = row["TYPTCA"].iloc[0]
            tn = row["TN"].iloc[0]

        typtca_Wert = 0
        if typtca == "1.2003.1EG":
            typtca_Wert = 1

        # Bestimmung des Zielwerts/Nennwerts: 100, 500 oder 1000
        if nennwert is not None:
            if int(nennwert) > 900:
                r0_sel_k = 1000.0
            elif 400 < int(nennwert) < 600:
                r0_sel_k = 500.0
            elif 90 < int(nennwert) < 110:
                r0_sel_k = 100.0
            else:
                r0_sel_k = int(nennwert)
        else:
            r0_sel_k = None

        # Verschiebung der Basiszeile nach unten für den nächsten Auftrag
        base_row = 2 + 8 * idx

        # Auftragsnummer im Sheet "DateiAus" prüfen
        eac_dateiaus = ws_dateiaus.cells(base_row, 2).value
        if eac_dateiaus != eac:
            print(f"Warnung: EAC in DateiAus ({eac_dateiaus}) passt nicht zu Auftragstabelle ({eac})")
            continue

        # Tabellenzeilen
        row_laser = base_row + 1
        row_nc_trimmen = base_row + 2
        row_ft_trimmen = base_row + 3
        row_nc_nachtrim = base_row + 4
        row_ft_nachtrim = base_row + 5
        row_ft_nachmessen = base_row + 6

        ws_dateiaus.cells(row_laser, 11).value = "Laser"
        ws_dateiaus.cells(row_nc_trimmen, 11).value = "NC_Trimmen"
        ws_dateiaus.cells(row_ft_trimmen, 11).value = "FT_Trimmen"
        ws_dateiaus.cells(row_nc_nachtrim, 11).value = "NC_Nachtrimmen"
        ws_dateiaus.cells(row_ft_nachtrim, 11).value = "FT_Nachtrimmen"
        ws_dateiaus.cells(row_ft_nachmessen, 11).value = "FT_Nachmessen"

        # Für Laser 1 bis 7
        for laser in range(1, 8):
            col = laser + 1
            ws_dateiaus.cells(row_laser, col + 10).value = laser

            # -------------------------
            # Trimmen
            # -------------------------
            ft_trim_value = ws_dateiaus.cells(row_ft_trimmen, col).value
            if ft_trim_value is not None:
                input_data = pd.DataFrame([{
                    "FT_T": ft_trim_value,
                    "ZIELWERT_T": ft_trim_value,
                    "NT_NM_NT/NM": ntnm_Wert,
                    "ZW_NC": nennwert,
                    "X_TEIL_ID": x_teil_id,
                    "R0_SEL_K": r0_sel_k
                }], columns=feature_cols_t)

                nc_trim_pred = round(float(rf_t.predict(input_data)[0]), 3)

                # Ergebnis in NC_Trimmen schreiben
                ws_dateiaus.cells(row_nc_trimmen, col + 10).value = float(nc_trim_pred)
                ws_dateiaus.cells(row_ft_trimmen, col + 10).value = ft_trim_value

            # -------------------------
            # Nachtrimmen
            # -------------------------
            ft_nachtrim_value = ws_dateiaus.cells(row_ft_nachtrim, col).value
            if ft_nachtrim_value is not None:
                input_data = pd.DataFrame([{
                    "FT_NT": ft_nachtrim_value,
                    "ZIELWERT_NT": ft_nachtrim_value,
                    "TN": tn,
                    "X_TEIL_ID": x_teil_id,
                    "R0_SEL_K": r0_sel_k,
                    "TYPTCA_1.2003.1EG": typtca_Wert
                }], columns=feature_cols_nt)

                nc_nachtrim_pred = round(float(rf_nt.predict(input_data)[0]), 3)

                # Ergebnis in NC_Nachtrimmen schreiben
                ws_dateiaus.cells(row_nc_nachtrim, col + 10).value = float(nc_nachtrim_pred)
                ws_dateiaus.cells(row_ft_nachtrim, col + 10).value = ft_nachtrim_value

            # -------------------------
            # Nachmessen
            # -------------------------
            ws_dateiaus.cells(row_ft_nachmessen, col + 10).value = ws_dateiaus.cells(row_ft_nachmessen, col).value

        # Progress Counter
        print(f"({idx + 1}|{anzahl_auftraege})")

    print("Berechnung fertig. Beende Programm...")
    time.sleep(1)

def run_elastic_net():
    print("Starte Elastic Net...")

    excel_path = "C:\\Users\\pythonProject\\NC_CALC_Programm.xlsm"
    try:
        wb = xw.Book.caller()
    except Exception:
        wb = xw.Book(excel_path)

    ws_auftraege = wb.sheets["Auftragstabelle"]
    ws_dateiaus = wb.sheets["DateiAus"]

    ws_dateiaus.cells(2, 21).value = "Elastic Net"

    # table_name = "Tabelle_Abfrage_von_OracleDB2025"
    feature_cols_t = ["FT_T", "X_TEIL_ID", "R0_SEL_K", "ZIELWERT_T"]
    feature_cols_nt = ["FT_NT", "X_TEIL_ID", "R0_SEL_K", "ZIELWERT_NT"]

    df = pd.read_excel(wb.fullname, sheet_name="Rohdaten")
    df.columns = df.columns.astype(str).str.strip()
    print("Daten gefunden. Beginne Modelltraining...")

    # Hilfsfunktionen
    def _to_numeric(s):
        return pd.to_numeric(s, errors="coerce")

    def _categorize_nennwert(v):
        """Erlaubte Kategorien: 100, 500, 1000. Wenn v exakt passt -> v.
        Sonst nächstgelegene Kategorie (nur wenn plausibel), ansonsten NaN."""
        try:
            if v is None or (isinstance(v, float) and np.isnan(v)):
                return np.nan
            vv = float(v)
        except Exception:
            return np.nan

        cats = np.array([100.0, 500.0, 1000.0])
        if vv in cats:
            return vv
        # nächste Kategorie, wenn nicht völlig daneben (Toleranz 25%)
        nearest = cats[np.argmin(np.abs(cats - vv))]
        if abs(nearest - vv) <= 0.25 * nearest:
            return float(nearest)
        return np.nan

    def _build_model():
        return Pipeline([
            ("scaler", StandardScaler()),
            ("enet", ElasticNetCV(
                alphas=np.logspace(-3, 1, 20),
                l1_ratio=[0.1, 0.2, 0.3, 0.4, 0.5],
                cv=5,
                max_iter=100000,
                n_jobs=-1,
                tol=1e-4
            ))
        ])

    def _select_training_data(df_clean, elastic_x_teil_id, elastic_nennwert_cat, zielwert_col, min_rows_same_id=50, min_rows_cat=80):
        """
        Auswahlstrategie:
        1) gleiche Nennwertkategorie + gleiche X_TEIL_ID (wenn genug Zeilen)
        2) sonst gleiche Nennwertkategorie (wenn genug)
        3) sonst Fallback: alle Zeilen aus df_clean
        """
        if df_clean is None or len(df_clean) == 0:
            return df_clean

        d = df_clean

        # Kategorie-Filter über Zielwert-Spalte (ZIELWERT_T / ZIELWERT_NT)
        if zielwert_col in d.columns and not (isinstance(elastic_nennwert_cat, float) and np.isnan(elastic_nennwert_cat)):
            d_cat = d[d[zielwert_col].apply(_categorize_nennwert) == float(elastic_nennwert_cat)]
        else:
            d_cat = d

        # X_TEIL_ID-Filter
        x_id_num = _to_numeric(pd.Series([elastic_x_teil_id])).iloc[0]
        if "X_TEIL_ID" in d_cat.columns and not (isinstance(x_id_num, float) and np.isnan(x_id_num)):
            d_cat_id = d_cat[d_cat["X_TEIL_ID"] == x_id_num]
        else:
            d_cat_id = d_cat

        if len(d_cat_id) >= min_rows_same_id:
            return d_cat_id
        if len(d_cat) >= min_rows_cat:
            return d_cat
        return d

    def _print_linear_importance(model, feature_cols, label, eac):
        try:
            enet = model.named_steps["enet"]
            coefs = np.asarray(enet.coef_, dtype=float)
            imp = pd.DataFrame({"Feature": feature_cols, "Importance": np.abs(coefs)})
            imp = imp.sort_values("Importance", ascending=False).reset_index(drop=True)
            print(f"[{eac}] {label} | Importance (skaliert):")
            print(imp.to_string(index=False))
        except Exception as ex:
            print(f"[{eac}] {label} | Importance konnte nicht berechnet werden: {ex}")

    # Filterungen
    mask = df["NC_CALC"] != "N"

    df_clean_t = create_clean_df(df, mask)
    df_clean_nt = df_clean_t.copy()

    df_clean_t = df_clean_t.dropna(subset=feature_cols_t)
    df_clean_nt = df_clean_nt.dropna(subset=feature_cols_nt)

    print(f"Zeilen eingelesen: {len(df)}")
    print(f"Zeilen nach Filterung Trimmen: {len(df_clean_t)}")
    print(f"Zeilen nach Filterung Nachtrimmen: {len(df_clean_nt)}")

    # Auftragstabelle
    auftraege = []
    row = 2
    while True:
        eac = ws_auftraege.cells(row, 2).value
        if eac is None:
            break
        x_teil_id = ws_auftraege.cells(row, 4).value
        nennwert = ws_auftraege.cells(row, 6).value

        auftraege.append({"EAC": eac, "X_TEIL_ID": x_teil_id, "Nennwert": nennwert})
        row += 1

    anzahl_auftraege = len(auftraege)
    print(f"Anzahl Aufträge gefunden: {len(auftraege)}")

    # Iteration durch jeden Auftrag in der Auftragstabelle
    for idx, auftrag in enumerate(auftraege):
        eac = auftrag["EAC"]
        x_teil_id = auftrag["X_TEIL_ID"]
        nennwert = auftrag["Nennwert"]

        nennwert_cat = _categorize_nennwert(nennwert)

        base_row = 2 + 8 * idx
        eac_dateiaus = ws_dateiaus.cells(base_row, 2).value
        if eac_dateiaus != eac:
            print(f"[ElasticNet] Warnung: EAC DateiAus ({eac_dateiaus}) != Auftragstabelle ({eac})")
            continue

        row_laser = base_row + 1
        row_nc_trimmen = base_row + 2
        row_ft_trimmen = base_row + 3
        row_nc_nachtrim = base_row + 4
        row_ft_nachtrim = base_row + 5
        row_ft_nachmessen = base_row + 6

        ws_dateiaus.cells(row_laser, 21).value = "Laser"
        ws_dateiaus.cells(row_nc_trimmen, 21).value = "NC_Trimmen"
        ws_dateiaus.cells(row_ft_trimmen, 21).value = "FT_Trimmen"
        ws_dateiaus.cells(row_nc_nachtrim, 21).value = "NC_Nachtrimmen"
        ws_dateiaus.cells(row_ft_nachtrim, 21).value = "FT_Nachtrimmen"
        ws_dateiaus.cells(row_ft_nachmessen, 21).value = "FT_Nachmessen"

        # Auftragsspezifische Trainingsdaten
        df_train_t = _select_training_data(
            df_clean=df_clean_t,
            elastic_x_teil_id=x_teil_id,
            elastic_nennwert_cat=nennwert_cat,
            zielwert_col="ZIELWERT_T",
            min_rows_same_id=50,
            min_rows_cat=80
        )
        df_train_nt = _select_training_data(
            df_clean=df_clean_nt,
            elastic_x_teil_id=x_teil_id,
            elastic_nennwert_cat=nennwert_cat,
            zielwert_col="ZIELWERT_NT",
            min_rows_same_id=50,
            min_rows_cat=80
        )

        # --- Modelle pro Auftrag trainieren (Trimmen & Nachtrimmen jeweils einmal) ---
        model_t = None
        model_nt = None

        # Trimmen trainieren
        if df_train_t is not None and len(df_train_t) >= 20 and all(c in df_train_t.columns for c in feature_cols_t) and "NC_T" in df_train_t.columns:
            X_t = df_train_t[feature_cols_t]
            y_t = df_train_t["NC_T"]

            model_t = _build_model()

            y_std = float(np.nanstd(y_t))
            y_nunique = int(pd.Series(y_t).nunique(dropna=True))

            print(f"[{eac}] y_t: min={y_t.min():.6f}, max={y_t.max():.6f}, std={y_t.std():.6e}, nunique={y_t.nunique()}")
            print(f"[{eac}] X_t std:\n{X_t.std().to_string()}")

            if y_nunique <= 5 or y_std < 1e-6:
                model_t = Pipeline([
                    ("scaler", StandardScaler()),
                    ("dummy", DummyRegressor(strategy="mean"))
                ])
                model_t.fit(X_t, y_t)
                print(f"[{eac}] Trimmen: y quasi konstant (std={y_std:.3e}, unique={y_nunique}) -> DummyRegressor(mean)")

            elif len(df_train_t) >= 50:
                X_tr, X_te, y_tr, y_te = train_test_split(X_t, y_t, test_size=0.2, random_state=42)
                model_t.fit(X_tr, y_tr)
                y_pred = model_t.predict(X_te)
                mae = mean_absolute_error(y_te, y_pred)
                rmse = np.sqrt(mean_squared_error(y_te, y_pred))
                r2 = r2_score(y_te, y_pred)
                print(f"[{eac}] Trimmen: TrainRows={len(df_train_t)} | MAE={mae:.4f} | RMSE={rmse:.4f} | R²={r2:.4f}")

            else:
                model_t.fit(X_t, y_t)
                print(f"[{eac}] Trimmen: TrainRows={len(df_train_t)} | (ohne Bewertungsmetriken)")

            # Feature Importance (Trimmen)
            _print_linear_importance(model_t, feature_cols_t, "Trimmen", eac)

        else:
            print(
                f"[{eac}] Trimmen: zu wenig/ungeeignete Trainingsdaten (Rows={len(df_train_nt) if df_train_nt is not None else 0})")

        # Nachtrimmen trainieren
        if df_train_nt is not None and len(df_train_nt) >= 20 and all(c in df_train_nt.columns for c in feature_cols_nt) and "NC_NT" in df_train_nt.columns:
            X_nt = df_train_nt[feature_cols_nt]
            y_nt = df_train_nt["NC_NT"]

            model_nt = _build_model()

            y_std = float(np.nanstd(y_nt))
            y_nunique = int(pd.Series(y_nt).nunique(dropna=True))

            if y_nunique <= 5 or y_std < 1e-3:
                model_nt = Pipeline([
                    ("scaler", StandardScaler()),
                    ("dummy", DummyRegressor(strategy="mean"))
                ])
                model_nt.fit(X_nt, y_nt)
                print(f"[{eac}] Trimmen: y quasi konstant (std={y_std:.3e}, unique={y_nunique}) -> DummyRegressor(mean)")

            elif len(df_train_nt) >= 50:
                X_tr, X_te, y_tr, y_te = train_test_split(X_nt, y_nt, test_size=0.2, random_state=42)
                model_nt.fit(X_tr, y_tr)
                y_pred = model_nt.predict(X_te)
                mae = mean_absolute_error(y_te, y_pred)
                rmse = np.sqrt(mean_squared_error(y_te, y_pred))
                r2 = r2_score(y_te, y_pred)
                print(f"[{eac}] Nachtrimmen: TrainRows={len(df_train_nt)} | MAE={mae:.4f} | RMSE={rmse:.4f} | R²={r2:.4f}")
            else:
                model_nt.fit(X_nt, y_nt)
                print(f"[{eac}] Nachtrimmen: TrainRows={len(df_train_nt)} | (ohne Bewertungsmetriken)")

            # Feature Importance (Nachtrimmen)
            _print_linear_importance(model_nt, feature_cols_nt, "Nachtrimmen", eac)

        else:
            print(
                f"[{eac}] Nachtrimmen: zu wenig/ungeeignete Trainingsdaten (Rows={len(df_train_nt) if df_train_nt is not None else 0})")

        for laser in range(1, 8):
            col = laser + 1
            out_col = col + 20

            ws_dateiaus.cells(row_laser, out_col).value = laser

            # -------------------------
            # Trimmen
            # -------------------------
            ft_trim_value = ws_dateiaus.cells(row_ft_trimmen, col).value
            if ft_trim_value is not None and model_t is not None:
                input_data = pd.DataFrame([{
                    "FT_T": float(ft_trim_value),
                    "ZIELWERT_T": float(ft_trim_value),
                    "X_TEIL_ID": _to_numeric(pd.Series([x_teil_id])).iloc[0],
                    "R0_SEL_K": float(nennwert) if nennwert is not None else np.nan

                    # Optionale Features
                    #"ZW_NC": float(nennwert) if nennwert is not None else np.nan,
                    #"LASER_T": int(laser),

                }], columns=feature_cols_t)

                pred = float(model_t.predict(input_data)[0])
                pred = round(pred, 3)

                ws_dateiaus.cells(row_nc_trimmen, out_col).value = pred
                ws_dateiaus.cells(row_ft_trimmen, out_col).value = ft_trim_value

            # -------------------------
            # Nachtrimmen
            # -------------------------
            ft_nachtrim_value = ws_dateiaus.cells(row_ft_nachtrim, col).value
            if ft_nachtrim_value is not None and model_nt is not None:
                input_data = pd.DataFrame([{
                    "FT_NT": float(ft_nachtrim_value),
                    "ZIELWERT_NT": float(ft_nachtrim_value),
                    "X_TEIL_ID": _to_numeric(pd.Series([x_teil_id])).iloc[0],
                    "R0_SEL_K": float(nennwert) if nennwert is not None else np.nan
                }], columns=feature_cols_nt)

                pred = float(model_nt.predict(input_data)[0])
                pred = round(pred, 3)

                ws_dateiaus.cells(row_nc_nachtrim, out_col).value = pred
                ws_dateiaus.cells(row_ft_nachtrim, out_col).value = ft_nachtrim_value

            # -------------------------
            # Nachmessen
            # -------------------------
            ws_dateiaus.cells(row_ft_nachmessen, col + 20).value = ws_dateiaus.cells(row_ft_nachmessen, col).value

        # Progress Counter
        print(f"({idx + 1}|{anzahl_auftraege})")

    print("Fertig. Beende Programm...")
    time.sleep(1)


def run_ridge():
    print("Starte Ridge Regression...")
    excel_path = "C:\\Users\\pythonProject\\NC_CALC_Programm.xlsm"
    try:
        wb = xw.Book.caller()
    except Exception:
        wb = xw.Book(excel_path)

    ws_auftraege = wb.sheets["Auftragstabelle"]
    ws_dateiaus = wb.sheets["DateiAus"]

    # Feature-Auswahl
    feature_cols = ["X_TEIL_ID", "ZW_NC", "R0_SEL_K"]
    feature_cols_t = ["FT_T", "LASER_T"]
    feature_cols_nt = ["FT_NT", "LASER_NT"]

    feature_cols_t = feature_cols_t + feature_cols
    feature_cols_nt = feature_cols_nt + feature_cols

    df = pd.read_excel(wb.fullname, sheet_name="Rohdaten")
    df.columns = df.columns.astype(str).str.strip()

    mask_t = (
            df["NC_T"].notna() &
            df["FT_T"].notna()
    )

    mask_nt = (
            df["NC_NT"].notna() &
            df["FT_NT"].notna()
    )

    df_clean_t = create_clean_df(df, mask_t)
    df_clean_nt = create_clean_df(df, mask_nt)

    df_clean_t = df_clean_t.dropna(subset=feature_cols_t)
    df_clean_nt = df_clean_nt.dropna(subset=feature_cols_nt)
    print("Daten gefunden. Beginne Modelltraining...")

    # Hilfsfunktionen
    def _to_numeric(s):
        return pd.to_numeric(s, errors="coerce")

    def _categorize_nennwert(v):
        """Erlaubte Kategorien: 100, 500, 1000. Wenn v exakt passt -> v.
        Sonst nächstgelegene Kategorie (nur wenn plausibel), ansonsten NaN."""
        try:
            if v is None or (isinstance(v, float) and np.isnan(v)):
                return np.nan
            vv = float(v)
        except Exception:
            return np.nan

        cats = np.array([100.0, 500.0, 1000.0])
        if vv in cats:
            return vv
        # nächste Kategorie, wenn nicht völlig daneben (Toleranz 25%)
        nearest = cats[np.argmin(np.abs(cats - vv))]
        if abs(nearest - vv) <= 0.25 * nearest:
            return float(nearest)
        return np.nan

    def _build_model():
        return Pipeline([
            ("scaler", StandardScaler()),
            ("ridge", RidgeCV(
                alphas=np.logspace(-3, 3, 20),
                cv=5
            ))
        ])

    def _select_training_data(df_clean, elastic_x_teil_id, elastic_nennwert_cat, zielwert_col, min_rows_same_id=50,
                              min_rows_cat=80):
        """
        Auswahlstrategie:
        1) gleiche Nennwertkategorie + gleiche X_TEIL_ID (wenn genug Zeilen)
        2) sonst gleiche Nennwertkategorie (wenn genug)
        3) sonst Fallback: alles df_clean
        """
        if df_clean is None or len(df_clean) == 0:
            return df_clean

        d = df_clean

        # Kategorie-Filter über Zielwert-Spalte (ZIELWERT_T / ZIELWERT_NT)
        if zielwert_col in d.columns and not (
                isinstance(elastic_nennwert_cat, float) and np.isnan(elastic_nennwert_cat)):
            d_cat = d[d[zielwert_col].apply(_categorize_nennwert) == float(elastic_nennwert_cat)]
        else:
            d_cat = d

        # X_TEIL_ID-Filter (numerisch)
        x_id_num = _to_numeric(pd.Series([elastic_x_teil_id])).iloc[0]
        if "X_TEIL_ID" in d_cat.columns and not (isinstance(x_id_num, float) and np.isnan(x_id_num)):
            d_cat_id = d_cat[d_cat["X_TEIL_ID"] == x_id_num]
        else:
            d_cat_id = d_cat

        if len(d_cat_id) >= min_rows_same_id:
            return d_cat_id
        if len(d_cat) >= min_rows_cat:
            return d_cat
        return d

    def _print_linear_importance(model, feature_cols, label, eac):
        try:
            enet = model.named_steps["enet"]
            coefs = np.asarray(enet.coef_, dtype=float)
            imp = pd.DataFrame({"Feature": feature_cols, "Importance": np.abs(coefs)})
            imp = imp.sort_values("Importance", ascending=False).reset_index(drop=True)
            print(f"[{eac}] {label} | Importance (skaliert):")
            print(imp.to_string(index=False))
        except Exception as ex:
            print(f"[{eac}] {label} | Importance konnte nicht berechnet werden: {ex}")

    # Auftragstabelle
    auftraege = []
    row = 2
    while True:
        eac = ws_auftraege.cells(row, 2).value
        if eac is None:
            break
        x_teil_id = ws_auftraege.cells(row, 4).value
        nennwert = ws_auftraege.cells(row, 6).value

        auftraege.append({"EAC": eac, "X_TEIL_ID": x_teil_id, "Nennwert": nennwert})
        row += 1

    anzahl_auftraege = len(auftraege)
    print(f"Anzahl Aufträge gefunden: {len(auftraege)}")

    ws_dateiaus.cells(2, 31).value = "Ridge Regression"

    # Iteration durch jeden Auftrag in der Auftragstabelle
    for idx, auftrag in enumerate(auftraege):
        eac = auftrag["EAC"]
        x_teil_id = auftrag["X_TEIL_ID"]
        nennwert = auftrag["Nennwert"]

        nennwert_cat = _categorize_nennwert(nennwert)

        base_row = 2 + 8 * idx
        eac_dateiaus = ws_dateiaus.cells(base_row, 2).value
        if eac_dateiaus != eac:
            print(f"[ElasticNet] Warnung: EAC DateiAus ({eac_dateiaus}) != Auftragstabelle ({eac})")
            continue

        row_laser = base_row + 1
        row_nc_trimmen = base_row + 2
        row_ft_trimmen = base_row + 3
        row_nc_nachtrim = base_row + 4
        row_ft_nachtrim = base_row + 5
        row_ft_nachmessen = base_row + 6

        ws_dateiaus.cells(row_laser, 31).value = "Laser"
        ws_dateiaus.cells(row_nc_trimmen, 31).value = "NC_Trimmen"
        ws_dateiaus.cells(row_ft_trimmen, 31).value = "FT_Trimmen"
        ws_dateiaus.cells(row_nc_nachtrim, 31).value = "NC_Nachtrimmen"
        ws_dateiaus.cells(row_ft_nachtrim, 31).value = "FT_Nachtrimmen"
        ws_dateiaus.cells(row_ft_nachmessen, 31).value = "FT_Nachmessen"

        # Auftragsspezifische Trainingsdaten
        df_train_t = _select_training_data(
            df_clean=df_clean_t,
            elastic_x_teil_id=x_teil_id,
            elastic_nennwert_cat=nennwert_cat,
            zielwert_col="ZIELWERT_T",
            min_rows_same_id=50,
            min_rows_cat=80
        )
        df_train_nt = _select_training_data(
            df_clean=df_clean_nt,
            elastic_x_teil_id=x_teil_id,
            elastic_nennwert_cat=nennwert_cat,
            zielwert_col="ZIELWERT_NT",
            min_rows_same_id=50,
            min_rows_cat=80
        )

        # --- Modelle pro Auftrag trainieren (Trimmen & Nachtrimmen jeweils einmal) ---
        model_t = None
        model_nt = None

        # Trimmen trainieren
        if df_train_t is not None and len(df_train_t) >= 20 and all(
                c in df_train_t.columns for c in feature_cols_t) and "NC_T" in df_train_t.columns:
            X_t = df_train_t[feature_cols_t]
            y_t = df_train_t["NC_T"]

            model_t = _build_model()

            y_std = float(np.nanstd(y_t))
            y_nunique = int(pd.Series(y_t).nunique(dropna=True))

            print(
                f"[{eac}] y_t: min={y_t.min():.6f}, max={y_t.max():.6f}, std={y_t.std():.6e}, nunique={y_t.nunique()}")
            print(f"[{eac}] X_t std:\n{X_t.std().to_string()}")

            if y_nunique <= 5 or y_std < 1e-6:
                model_t = Pipeline([
                    ("scaler", StandardScaler()),
                    ("dummy", DummyRegressor(strategy="mean"))
                ])
                model_t.fit(X_t, y_t)
                print(
                    f"[{eac}] Trimmen: y quasi konstant (std={y_std:.3e}, unique={y_nunique}) -> DummyRegressor(mean)")

            elif len(df_train_t) >= 50:
                X_tr, X_te, y_tr, y_te = train_test_split(X_t, y_t, test_size=0.2, random_state=42)
                model_t.fit(X_tr, y_tr)
                y_pred = model_t.predict(X_te)
                mae = mean_absolute_error(y_te, y_pred)

                r2 = r2_score(y_te, y_pred)
                print(f"[{eac}] Trimmen: TrainRows={len(df_train_t)} | MAE={mae:.4f} | R²={r2:.4f}")

            else:
                model_t.fit(X_t, y_t)
                print(f"[{eac}] Trimmen: TrainRows={len(df_train_t)} | (ohne Bewertungsmetriken)")

            # Feature Importance (Trimmen)
            # _print_linear_importance(model_t, feature_cols_t, "Trimmen", eac)

        else:
            print(
                f"[{eac}] Trimmen: zu wenig/ungeeignete Trainingsdaten (Rows={len(df_train_nt) if df_train_nt is not None else 0})")

        # Nachtrimmen trainieren
        if df_train_nt is not None and len(df_train_nt) >= 20 and all(
                c in df_train_nt.columns for c in feature_cols_nt) and "NC_NT" in df_train_nt.columns:
            X_nt = df_train_nt[feature_cols_nt]
            y_nt = df_train_nt["NC_NT"]

            model_nt = _build_model()

            y_std = float(np.nanstd(y_nt))
            y_nunique = int(pd.Series(y_nt).nunique(dropna=True))

            if y_nunique <= 5 or y_std < 1e-3:
                model_nt = Pipeline([
                    ("scaler", StandardScaler()),
                    ("dummy", DummyRegressor(strategy="mean"))
                ])
                model_nt.fit(X_nt, y_nt)
                print(
                    f"[{eac}] Trimmen: y quasi konstant (std={y_std:.3e}, unique={y_nunique}) -> DummyRegressor(mean)")

            elif len(df_train_nt) >= 50:
                X_tr, X_te, y_tr, y_te = train_test_split(X_nt, y_nt, test_size=0.2, random_state=42)
                model_nt.fit(X_tr, y_tr)
                y_pred = model_nt.predict(X_te)
                mae = mean_absolute_error(y_te, y_pred)
                r2 = r2_score(y_te, y_pred)
                print(f"[{eac}] Nachtrimmen: TrainRows={len(df_train_nt)} | MAE={mae:.4f} | R²={r2:.4f}")
            else:
                model_nt.fit(X_nt, y_nt)
                print(f"[{eac}] Nachtrimmen: TrainRows={len(df_train_nt)} | (ohne Bewertungsmetriken)")

            # Feature Importance (Nachtrimmen)
            # _print_linear_importance(model_nt, feature_cols_nt, "Nachtrimmen", eac)

        else:
            print(
                f"[{eac}] Nachtrimmen: zu wenig/ungeeignete Trainingsdaten (Rows={len(df_train_nt) if df_train_nt is not None else 0})")

        for laser in range(1, 8):
            col = laser + 1     # Spalte 2 bis 9
            out_col = col + 30  # Spalte 21 bis 28

            ws_dateiaus.cells(row_laser, out_col).value = laser

            # -------------------------
            # Trimmen
            # -------------------------
            ft_trim_value = ws_dateiaus.cells(row_ft_trimmen, col).value
            if ft_trim_value is not None and model_t is not None:
                input_data = pd.DataFrame([{
                    "FT_T": float(ft_trim_value),
                    "ZW_NC": float(nennwert) if nennwert is not None else np.nan,
                    "LASER_T": int(laser),
                    "X_TEIL_ID": _to_numeric(pd.Series([x_teil_id])).iloc[0],
                    "R0_SEL_K": float(nennwert) if nennwert is not None else np.nan
                }], columns=feature_cols_t)

                pred = float(model_t.predict(input_data)[0])
                pred = round(pred, 3)

                ws_dateiaus.cells(row_nc_trimmen, out_col).value = pred
                ws_dateiaus.cells(row_ft_trimmen, out_col).value = ft_trim_value

            # -------------------------
            # Nachtrimmen
            # -------------------------
            ft_nachtrim_value = ws_dateiaus.cells(row_ft_nachtrim, col).value
            if ft_nachtrim_value is not None and model_nt is not None:
                input_data = pd.DataFrame([{
                    "FT_NT": float(ft_nachtrim_value),
                    "ZW_NC": float(nennwert) if nennwert is not None else np.nan,
                    "LASER_NT": int(laser),
                    "X_TEIL_ID": _to_numeric(pd.Series([x_teil_id])).iloc[0],
                    "R0_SEL_K": float(nennwert) if nennwert is not None else np.nan
                }], columns=feature_cols_nt)

                pred = float(model_nt.predict(input_data)[0])
                pred = round(pred, 3)

                ws_dateiaus.cells(row_nc_nachtrim, out_col).value = pred
                ws_dateiaus.cells(row_ft_nachtrim, out_col).value = ft_nachtrim_value

            # -------------------------
            # Nachmessen
            # -------------------------
            ws_dateiaus.cells(row_ft_nachmessen, col + 30).value = ws_dateiaus.cells(row_ft_nachmessen, col).value

        # Progress Counter
        print(f"({idx + 1}|{anzahl_auftraege})")

    print("Fertig. Beende Programm...")
    time.sleep(1)


def run_gaussian_process_regression_linear():
    print("Starte Gaussian Process Regression...")

    excel_path = "C:\\Users\\pythonProject\\NC_CALC_Programm.xlsm"
    try:
        wb = xw.Book.caller()
    except Exception:
        wb = xw.Book(excel_path)

    ws_auftraege = wb.sheets["Auftragstabelle"]
    ws_dateiaus = wb.sheets["DateiAus"]

    ws_dateiaus.cells(2, 31).value = "Gaussian Process Regression"

    feature_cols_t = ["ZIELWERT_T", "ZW_NC", "LASER_T", "X_TEIL_ID", "R0_SEL_K"]
    feature_cols_nt = ["ZIELWERT_NT", "ZW_NC", "LASER_NT", "X_TEIL_ID", "R0_SEL_K"]

    df = pd.read_excel(wb.fullname, sheet_name="Rohdaten")
    df.columns = df.columns.astype(str).str.strip()
    print("Daten gefunden. Beginne Modelltraining...")

    # -------------------------
    # Hilfsfunktionen
    # -------------------------
    def _to_numeric(s):
        return pd.to_numeric(s, errors="coerce")

    def _categorize_nennwert(v):
        """Erlaubte Kategorien: 67.6, 100, 500, 1000. Wenn v exakt passt -> v.
        67.6 ist eine Extrawurst und kann bei der Berechnung Convergence Warnungen erzeugen.
        Sonst nächstgelegene Kategorie (nur wenn plausibel), ansonsten NaN."""
        try:
            if v is None or (isinstance(v, float) and np.isnan(v)):
                return np.nan
            vv = float(v)
        except Exception:
            return np.nan

        cats = np.array([67.6, 100.0, 500.0, 1000.0])
        if vv in cats:
            return vv

        nearest = cats[np.argmin(np.abs(cats - vv))]
        if abs(nearest - vv) <= 0.2 * nearest:
            return float(nearest)
        return np.nan

    def _build_gpr_model():
        kernel = (
                ConstantKernel(1.0, (1e-3, 1e4))
                * DotProduct(sigma_0=1.0, sigma_0_bounds=(1e-8, 1e3))
                + ConstantKernel(1.0, (1e-3, 1e3))
                * RBF(length_scale=1.0, length_scale_bounds=(1e-4, 1e3))
                + WhiteKernel(noise_level=1e-2, noise_level_bounds=(1e-12, 1e1))
        )
        gpr = GaussianProcessRegressor(
            kernel=kernel,
            alpha=1e-5,
            normalize_y=True,
            n_restarts_optimizer=5,
            random_state=42
        )
        return Pipeline([("scaler", StandardScaler()), ("gpr", gpr)])

    def build_residual_gpr():
        kernel = (
                ConstantKernel(1.0, (1e-3, 1e4))
                * RBF(length_scale=1.0, length_scale_bounds=(1e-2, 1e3))
                + WhiteKernel(noise_level=1e-3, noise_level_bounds=(1e-12, 1e1))
        )
        gpr = GaussianProcessRegressor(
            kernel=kernel,
            alpha=1e-5,
            normalize_y=True,
            n_restarts_optimizer=3,
            random_state=42
        )
        return Pipeline([("scaler", StandardScaler()), ("gpr", gpr)])

    def _select_training_data(df_clean, elastic_x_teil_id, elastic_nennwert_cat, zielwert_col, min_rows_same_id=50, min_rows_cat=80):
        if df_clean is None or len(df_clean) == 0:
            return df_clean

        d = df_clean

        if zielwert_col in d.columns and not (isinstance(elastic_nennwert_cat, float) and np.isnan(elastic_nennwert_cat)):
            d_cat = d[d[zielwert_col].apply(_categorize_nennwert) == float(elastic_nennwert_cat)]
        else:
            d_cat = d

        x_id_num = _to_numeric(pd.Series([elastic_x_teil_id])).iloc[0]
        if "X_TEIL_ID" in d_cat.columns and not (isinstance(x_id_num, float) and np.isnan(x_id_num)):
            d_cat_id = d_cat[d_cat["X_TEIL_ID"] == x_id_num]
        else:
            d_cat_id = d_cat

        if len(d_cat_id) >= min_rows_same_id:
            return d_cat_id
        if len(d_cat) >= min_rows_cat:
            return d_cat
        return d

    def slope_from_nennwert_cat(nennwert_cat: float) -> float:
        # Nennwert Kategorien → Steigung
        if nennwert_cat == 67.6:
            return 1.3
        if nennwert_cat == 100.0:
            return 1.0
        if nennwert_cat == 500.0:
            return 0.2
        if nennwert_cat == 1000.0:
            return 0.1
        return np.nan

    def estimate_intercept(ft: pd.Series, y: pd.Series, slope: float, default: float = 0.0) -> float:
        ft_num = pd.to_numeric(ft, errors="coerce")
        y_num = pd.to_numeric(y, errors="coerce")
        resid = y_num - slope * ft_num
        b = float(np.nanmean(resid))
        if np.isnan(b) or np.isinf(b):
            return float(default)
        return b

    mask_t = df["NC_CALC"] != "N"

    df_clean_t = create_clean_df(df, mask_t)
    df_clean_nt = df_clean_t.copy()

    df_clean_t = df_clean_t.dropna(subset=feature_cols_t)
    df_clean_nt = df_clean_nt.dropna(subset=feature_cols_nt)

    print(f"Zeilen eingelesen: {len(df)}")
    print(f"Zeilen nach Filterung Trimmen: {len(df_clean_t)}")
    print(f"Zeilen nach Filterung Nachtrimmen: {len(df_clean_nt)}")

    # -------------------------
    # Auftragstabelle
    # -------------------------
    auftraege = []
    row = 2
    while True:
        try:
            eac = int(ws_auftraege.cells(row, 2).value)
        except:
            eac = ws_auftraege.cells(row, 2).value

        if eac is None:
            break
        x_teil_id = ws_auftraege.cells(row, 4).value
        nennwert = ws_auftraege.cells(row, 6).value

        auftraege.append({"EAC": eac, "X_TEIL_ID": x_teil_id, "Nennwert": nennwert})
        row += 1

    anzahl_auftraege = len(auftraege)
    print(f"Anzahl Aufträge gefunden: {anzahl_auftraege}")

    # -------------------------
    # Iteration über Aufträge
    # -------------------------
    for idx, auftrag in enumerate(auftraege):
        eac = auftrag["EAC"]
        x_teil_id = auftrag["X_TEIL_ID"]
        nennwert = auftrag["Nennwert"]

        nennwert_cat = _categorize_nennwert(nennwert)

        base_row = 2 + 8 * idx
        eac_dateiaus = ws_dateiaus.cells(base_row, 2).value
        if eac_dateiaus != eac:
            print(f"[GPR] Warnung: EAC DateiAus ({eac_dateiaus}) != Auftragstabelle ({eac})")
            continue

        row_laser = base_row + 1
        row_nc_trimmen = base_row + 2
        row_ft_trimmen = base_row + 3
        row_nc_nachtrim = base_row + 4
        row_ft_nachtrim = base_row + 5
        row_ft_nachmessen = base_row + 6

        ws_dateiaus.cells(row_laser, 31).value = "Laser"
        ws_dateiaus.cells(row_nc_trimmen, 31).value = "NC_Trimmen"
        ws_dateiaus.cells(row_ft_trimmen, 31).value = "FT_Trimmen"
        ws_dateiaus.cells(row_nc_nachtrim, 31).value = "NC_Nachtrimmen"
        ws_dateiaus.cells(row_ft_nachtrim, 31).value = "FT_Nachtrimmen"
        ws_dateiaus.cells(row_ft_nachmessen, 31).value = "FT_Nachmessen"

        # Auftragsspezifische Trainingsdaten
        df_train_t = _select_training_data(
            df_clean=df_clean_t,
            elastic_x_teil_id=x_teil_id,
            elastic_nennwert_cat=nennwert_cat,
            zielwert_col="ZIELWERT_T",
            min_rows_same_id=50,
            min_rows_cat=80
        )
        df_train_nt = _select_training_data(
            df_clean=df_clean_nt,
            elastic_x_teil_id=x_teil_id,
            elastic_nennwert_cat=nennwert_cat,
            zielwert_col="ZIELWERT_NT",
            min_rows_same_id=50,
            min_rows_cat=80
        )

        # --- Modelle pro Auftrag trainieren (Trimmen & Nachtrimmen) ---
        use_trend_t = False
        use_trend_nt = False

        slope_t = 0.0
        intercept_t = 0.0

        slope_nt = 0.0
        intercept_nt = 0.0

        # Trend aktivieren und bestimmen, falls Nennwertkategorie bekannt
        if not (isinstance(nennwert_cat, float) and np.isnan(nennwert_cat)):
            s = slope_from_nennwert_cat(float(nennwert_cat))
            if not np.isnan(s):
                slope_t = float(s)
                slope_nt = float(s)
                use_trend_t = True
                use_trend_nt = True

        # -------------------------
        # Trimmen trainieren
        # -------------------------
        if (
            df_train_t is not None
            and len(df_train_t) >= 20
            and all(c in df_train_t.columns for c in feature_cols_t)
            and "NC_T" in df_train_t.columns
        ):
            X_t = df_train_t[feature_cols_t]
            y_t = df_train_t["NC_T"]

            y_std = float(np.nanstd(y_t))
            y_nunique = int(pd.Series(y_t).nunique(dropna=True))

            print(f"[{eac}] y_t: min={y_t.min():.6f}, max={y_t.max():.6f}, std={y_t.std():.6e}, nunique={y_t.nunique()}")
            print(f"[{eac}] X_t std:\n{X_t.std().to_string()}")

            # Fallback bei quasi-konstantem Ziel → Berechnung des NC-Durchschnittswerts
            if y_nunique <= 5 or y_std < 1e-6:
                model_t = Pipeline([
                    ("scaler", StandardScaler()),
                    ("dummy", DummyRegressor(strategy="mean"))
                ])
                model_t.fit(X_t, y_t)
                slope_t, intercept_t = 0.0, 0.0
                print(f"[{eac}] Trimmen: y quasi konstant (std={y_std:.3e}, unique={y_nunique}) -> DummyRegressor(mean)")
            else:
                if use_trend_t:
                    # Intercept schätzen
                    intercept_t = estimate_intercept(df_train_t["FT_T"], y_t, slope_t, default=float(np.nanmean(y_t)))

                    # Residuen
                    ft_series = pd.to_numeric(df_train_t["FT_T"], errors="coerce")
                    y_res = y_t - (slope_t * ft_series + intercept_t)

                    y_res_std = float(np.nanstd(y_res))
                    y_res_nunique = int(pd.Series(y_res).nunique(dropna=True))

                    # Wenn Residuen extrem klein → nur Trend
                    if y_res_nunique <= 5 or y_res_std < 1e-6:
                        # GP weglassen, nur Trend nutzen
                        model_t = None
                        print(f"[{eac}] Trimmen: Trend dominiert (Residual std={y_res_std:.2e}) -> nur Trend, kein GP")
                    else:
                        # GPR-Modell mit linearem Trend
                        model_t = build_residual_gpr()
                        if len(df_train_t) >= 50:
                            X_tr, X_te, y_tr, y_te = train_test_split(X_t, y_res, test_size=0.2, random_state=42)
                            model_t.fit(X_tr, y_tr)
                            r_pred = model_t.predict(X_te)
                            mae = mean_absolute_error(y_te, r_pred)
                            rmse = np.sqrt(mean_squared_error(y_te, r_pred))
                            r2 = r2_score(y_te, r_pred)
                            print(
                                f"[{eac}] Trimmen (Residual-GP): TrainRows={len(df_train_t)} | MAE(res)={mae:.4f} | RMSE={rmse:.4f} | R²(res)={r2:.4f}")
                        else:
                            model_t.fit(X_t, y_res)
                            print(
                                f"[{eac}] Trimmen (Residual-GP): TrainRows={len(df_train_t)} | (ohne Bewertungsmetriken)")
                        print(f"[{eac}] Trimmen Trend: slope={slope_t}, intercept={intercept_t:.6f}")
                else:
                    # Kein Trend: klassisches GPR mit Trend_Kernel
                    model_t = _build_gpr_model()
                    if len(df_train_t) >= 50:
                        X_tr, X_te, y_tr, y_te = train_test_split(X_t, y_t, test_size=0.2, random_state=42)
                        model_t.fit(X_tr, y_tr)
                        y_pred = model_t.predict(X_te)
                        mae = mean_absolute_error(y_te, y_pred)
                        rmse = np.sqrt(mean_squared_error(y_te, y_pred))
                        r2 = r2_score(y_te, y_pred)
                        print(
                            f"[{eac}] Trimmen (GPR direkt): TrainRows={len(df_train_t)} | MAE={mae:.4f} | RMSE={rmse:.4f} | R²={r2:.4f}")
                    else:
                        model_t.fit(X_t, y_t)
                        print(f"[{eac}] Trimmen (GPR direkt): TrainRows={len(df_train_t)} | (ohne Bewertungsmetriken)")
        else:
            print(f"[{eac}] Trimmen: zu wenig/ungeeignete Trainingsdaten (Rows={len(df_train_t) if df_train_t is not None else 0})")
            model_t = None
            slope_t, intercept_t = 0.0, 0.0

        # -------------------------
        # Nachtrimmen trainieren
        # -------------------------
        if (
            df_train_nt is not None
            and len(df_train_nt) >= 20
            and all(c in df_train_nt.columns for c in feature_cols_nt)
            and "NC_NT" in df_train_nt.columns
        ):
            X_nt = df_train_nt[feature_cols_nt]
            y_nt = df_train_nt["NC_NT"]

            y_std = float(np.nanstd(y_nt))
            y_nunique = int(pd.Series(y_nt).nunique(dropna=True))

            if y_nunique <= 5 or y_std < 1e-3:
                model_nt = Pipeline([
                    ("scaler", StandardScaler()),
                    ("dummy", DummyRegressor(strategy="mean"))
                ])
                model_nt.fit(X_nt, y_nt)
                print(f"[{eac}] Nachtrimmen: y quasi konstant (std={y_std:.3e}, unique={y_nunique}) -> DummyRegressor(mean)")
            else:

                if use_trend_nt:
                    intercept_nt = estimate_intercept(df_train_nt["FT_NT"], y_nt, slope_nt,
                                                           default=float(np.nanmean(y_nt)))

                    ft_series = pd.to_numeric(df_train_nt["FT_NT"], errors="coerce")
                    y_res = y_nt - (slope_nt * ft_series + intercept_nt)

                    y_res_std = float(np.nanstd(y_res))
                    y_res_nunique = int(pd.Series(y_res).nunique(dropna=True))

                    # Nur Trend, kein GP
                    if y_res_nunique <= 5 or y_res_std < 1e-6:
                        model_nt = None
                        print(
                            f"[{eac}] Nachtrimmen: Trend dominiert (Residual std={y_res_std:.2e}) -> nur Trend, kein GP")
                    else:
                        # GPR-Modell mit linearem Trend
                        model_nt = build_residual_gpr()
                        if len(df_train_nt) >= 50:
                            X_tr, X_te, y_tr, y_te = train_test_split(X_nt, y_res, test_size=0.2, random_state=42)
                            model_nt.fit(X_tr, y_tr)
                            y_pred = model_nt.predict(X_te)
                            mae = mean_absolute_error(y_te, y_pred)
                            rmse = np.sqrt(mean_squared_error(y_te, y_pred))
                            r2 = r2_score(y_te, y_pred)
                            print(
                                f"[{eac}] Nachtrimmen (Residual-GP): TrainRows={len(df_train_nt)} | MAE(res)={mae:.4f} | RMSE={rmse:.4f} | R²(res)={r2:.4f}")
                        else:
                            model_nt.fit(X_nt, y_res)
                            print(
                                f"[{eac}] Nachtrimmen (Residual-GP): TrainRows={len(df_train_nt)} | (ohne Bewertungsmetriken)")
                        print(f"[{eac}] Nachtrimmen Trend: slope={slope_nt}, intercept={intercept_nt:.6f}")
                else:
                    # Fallback: Normales GPR-Modell mit Trend_Kernel
                    model_nt = _build_gpr_model()
                    if len(df_train_nt) >= 50:
                        X_tr, X_te, y_tr, y_te = train_test_split(X_nt, y_nt, test_size=0.2, random_state=42)
                        model_nt.fit(X_tr, y_tr)
                        y_pred = model_nt.predict(X_te)
                        mae = mean_absolute_error(y_te, y_pred)
                        rmse = np.sqrt(mean_squared_error(y_te, y_pred))
                        r2 = r2_score(y_te, y_pred)
                        print(
                            f"[{eac}] Nachtrimmen (GPR direkt): TrainRows={len(df_train_nt)} | MAE={mae:.4f} | RMSE={rmse:.4f} | R²={r2:.4f}")
                    else:
                        model_nt.fit(X_nt, y_nt)
                        print(
                            f"[{eac}] Nachtrimmen (GPR direkt): TrainRows={len(df_train_nt)} | (ohne Bewertungsmetriken)")

        else:
            print(f"[{eac}] Nachtrimmen: zu wenig/ungeeignete Trainingsdaten (Rows={len(df_train_nt) if df_train_nt is not None else 0})")
            model_nt = None
            slope_nt, intercept_nt = 0.0, 0.0

        # -------------------------
        # Vorhersagen schreiben (Laser 1..7)
        # -------------------------
        for laser in range(1, 8):
            col = laser + 1      # Spalte 2 bis 9
            out_col = col + 30   # Spalte 21 bis 28

            ws_dateiaus.cells(row_laser, out_col).value = laser

            # --------------------
            # Trimmen
            # --------------------
            ft_trim_value = ws_dateiaus.cells(row_ft_trimmen, col).value
            if ft_trim_value is not None:
                ft = float(ft_trim_value)

                input_data = pd.DataFrame([{
                    "ZIELWERT_T": ft,
                    "ZW_NC": float(nennwert) if nennwert is not None else np.nan,
                    "LASER_T": int(laser),
                    "X_TEIL_ID": _to_numeric(pd.Series([x_teil_id])).iloc[0],
                    "R0_SEL_K": float(nennwert) if nennwert is not None else np.nan
                }], columns=feature_cols_t)

                """ Drei Fälle:
                A) Trend + Residual-GP vorhanden → Trend + Residuum
                B) Trend aktiv aber kein GP (Trend dominiert) → nur Trend
                C) Trend nicht aktiv → Modell direkt auf y (model_t muss existieren)
                """
                pred = None

                if use_trend_t:
                    # Trend-Teil definiert
                    base = slope_t * ft + float(intercept_t)
                    if model_t is None:
                        pred = base
                    else:
                        r_hat = float(model_t.predict(input_data)[0])
                        pred = base + r_hat
                else:
                    if model_t is not None:
                        pred = float(model_t.predict(input_data)[0])

                if pred is not None:
                    pred = round(float(pred), 3)
                    ws_dateiaus.cells(row_nc_trimmen, out_col).value = pred
                    ws_dateiaus.cells(row_ft_trimmen, out_col).value = ft_trim_value

            # --------------------
            # Nachtrimmen
            # --------------------
            ft_nachtrim_value = ws_dateiaus.cells(row_ft_nachtrim, col).value
            if ft_nachtrim_value is not None:
                ft = float(ft_nachtrim_value)

                input_data = pd.DataFrame([{
                    "ZIELWERT_NT": ft,
                    "ZW_NC": float(nennwert) if nennwert is not None else np.nan,
                    "LASER_NT": int(laser),
                    "X_TEIL_ID": _to_numeric(pd.Series([x_teil_id])).iloc[0],
                    "R0_SEL_K": float(nennwert) if nennwert is not None else np.nan
                }], columns=feature_cols_nt)

                pred = None

                if use_trend_nt:
                    base = slope_nt * ft + float(intercept_nt)
                    if model_nt is None:
                        pred = base
                    else:
                        r_hat = float(model_nt.predict(input_data)[0])
                        pred = base + r_hat
                else:
                    if model_nt is not None:
                        pred = float(model_nt.predict(input_data)[0])

                if pred is not None:
                    pred = round(float(pred), 3)
                    ws_dateiaus.cells(row_nc_nachtrim, out_col).value = pred
                    ws_dateiaus.cells(row_ft_nachtrim, out_col).value = ft_nachtrim_value

            # Nachmessen
            ws_dateiaus.cells(row_ft_nachmessen, col + 30).value = ws_dateiaus.cells(row_ft_nachmessen, col).value

        print(f"({idx + 1}|{anzahl_auftraege})")

    print("Fertig. Beende Programm...")
    time.sleep(1)

if __name__ == '__main__':
    print("Pycharm-Test:")
    # select_features()
    run_model()
    run_elastic_net()
    # run_ridge()
    run_gaussian_process_regression_linear()
