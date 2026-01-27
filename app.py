import pandas as pd
import streamlit as st
import io

st.title("Transfer Planning Tool")

# Load Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        data = pd.read_excel(uploaded_file)
        st.success("Excel file loaded successfully!")
        st.dataframe(data.head())

        if st.button("Generate Transfer List"):
            df = data.copy()

            # --- Required columns check ---
            required_cols = ["Bölge Müdürü", "Depo Kodu", "Madde Kodu", "İhtiyaç", "Transfer Edilebilir"]
            missing = [c for c in required_cols if c not in df.columns]
            if missing:
                st.error(f"Eksik kolonlar: {missing}")
                st.stop()

            # --- Normalize key columns to avoid Excel dtype issues ---
            df["Depo Kodu"] = df["Depo Kodu"].astype(str)
            df["Madde Kodu"] = df["Madde Kodu"].astype(str)

            # --- Clean numeric columns ---
            df["İhtiyaç"] = pd.to_numeric(df["İhtiyaç"], errors="coerce").fillna(0)
            df["Transfer Edilebilir"] = pd.to_numeric(df["Transfer Edilebilir"], errors="coerce").fillna(0)

            # Optional: clip negatives
            df.loc[df["İhtiyaç"] < 0, "İhtiyaç"] = 0
            df.loc[df["Transfer Edilebilir"] < 0, "Transfer Edilebilir"] = 0

            # --- Global state dicts (shared across both stages) ---
            need = df.set_index(["Depo Kodu", "Madde Kodu"])["İhtiyaç"].to_dict()
            availability = df.set_index(["Depo Kodu", "Madde Kodu"])["Transfer Edilebilir"].to_dict()

            transfer_list = []

            def apply_transfer(sender_depot, receiver_depot, item_code, qty, transfer_tipi):
                """Write a transfer row and update global state."""
                if qty <= 0:
                    return
                transfer_list.append({
                    "Transfer Tipi": transfer_tipi,   # ✅ Flag: Bölge içi / Bölge dışı
                    "Gönderen Depo": sender_depot,
                    "Alan Depo": receiver_depot,
                    "Madde Kodu": item_code,
                    "Transfer Miktarı": qty
                })
                availability[(sender_depot, item_code)] -= qty
                need[(receiver_depot, item_code)] -= qty

            # =========================================================
            # STAGE 1: Region (Bölge Müdürü) internal transfers
            # =========================================================
            for manager, group in df.groupby("Bölge Müdürü"):
                items_in_region = group["Madde Kodu"].unique()

                for item_code in items_in_region:
                    # Receivers: depots with remaining need for this item
                    receivers = group[group["Madde Kodu"] == item_code][["Depo Kodu"]].drop_duplicates()
                    receivers["need"] = receivers["Depo Kodu"].apply(lambda d: need.get((d, item_code), 0))
                    receivers = receivers[receivers["need"] > 0].sort_values("need", ascending=False)
                    if receivers.empty:
                        continue

                    # Senders: depots with remaining availability for this item
                    senders = group[group["Madde Kodu"] == item_code][["Depo Kodu"]].drop_duplicates()
                    senders["avail"] = senders["Depo Kodu"].apply(lambda d: availability.get((d, item_code), 0))
                    senders = senders[senders["avail"] > 0].sort_values("avail", ascending=False)
                    if senders.empty:
                        continue

                    # Greedy matching: highest need first, supply from highest availability
                    for _, r in receivers.iterrows():
                        recv = r["Depo Kodu"]
                        r_need = need.get((recv, item_code), 0)
                        if r_need <= 0:
                            continue

                        for _, s in senders.iterrows():
                            send = s["Depo Kodu"]
                            if send == recv:
                                continue

                            s_avail = availability.get((send, item_code), 0)
                            if s_avail <= 0 or r_need <= 0:
                                continue

                            qty = min(r_need, s_avail)
                            apply_transfer(send, recv, item_code, qty, "Bölge içi")
                            r_need -= qty

            # =========================================================
            # STAGE 2: Cross-region transfers (no region constraint)
            # =========================================================
            all_items = df["Madde Kodu"].unique()
            all_depots = df[["Depo Kodu"]].drop_duplicates()

            for item_code in all_items:
                # Global receivers
                receivers = all_depots.copy()
                receivers["need"] = receivers["Depo Kodu"].apply(lambda d: need.get((d, item_code), 0))
                receivers = receivers[receivers["need"] > 0].sort_values("need", ascending=False)
                if receivers.empty:
                    continue

                # Global senders
                senders = all_depots.copy()
                senders["avail"] = senders["Depo Kodu"].apply(lambda d: availability.get((d, item_code), 0))
                senders = senders[senders["avail"] > 0].sort_values("avail", ascending=False)
                if senders.empty:
                    continue

                for _, r in receivers.iterrows():
                    recv = r["Depo Kodu"]
                    r_need = need.get((recv, item_code), 0)
                    if r_need <= 0:
                        continue

                    for _, s in senders.iterrows():
                        send = s["Depo Kodu"]
                        if send == recv:
                            continue

                        s_avail = availability.get((send, item_code), 0)
                        if s_avail <= 0 or r_need <= 0:
                            continue

                        qty = min(r_need, s_avail)
                        apply_transfer(send, recv, item_code, qty, "Bölge dışı")
                        r_need -= qty

            transfer_df = pd.DataFrame(transfer_list)

            st.success("Transfer list generated successfully!")
            st.dataframe(transfer_df)

            # Download as Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                transfer_df.to_excel(writer, index=False)

            st.download_button(
                label="Download Transfer List",
                data=buffer.getvalue(),
                file_name="Transfer_Listesi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"Failed to load file: {e}")
