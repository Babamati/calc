import streamlit as st
from datetime import datetime, date
from decimal import Decimal
from dateutil.relativedelta import relativedelta
from fpdf import FPDF
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import math
import calendar
import tempfile
import os

st.set_page_config(page_title="Bond Calculator", layout="wide")
page_icon="🧮",

# =========================
# Financial helper methods
# =========================
def add_months(source_date, months):
    return source_date + relativedelta(months=months)


def compute_macaulay_duration(cashflows, discount_factors, times, price):
    if price == 0:
        return 0.0
    return sum(t * cf * df for t, cf, df in zip(times, cashflows, discount_factors)) / price


def compute_convexity(cashflows, times, ytm, price, freq, is_bill):
    if price == 0:
        return 0.0

    if is_bill:
        t = times[0]
        return (t * (t + 1)) / ((1 + ytm) ** 2)
    else:
        y_period = ytm / freq
        convexity_sum = sum(
            cf * t * (t + 1) / (1 + y_period) ** (t + 2)
            for cf, t in zip(cashflows, [t * freq for t in times])
        )
        return convexity_sum / (price * freq ** 2)


def get_last_coupon_date(settlement, freq, maturity):
    delta_months = 12 // freq
    anchor_day = maturity.day
    last_coupon = maturity

    while last_coupon > settlement:
        last_coupon = add_months(last_coupon, -delta_months)
        try:
            last_coupon = last_coupon.replace(day=anchor_day)
        except ValueError:
            last_coupon = (
                last_coupon.replace(day=1)
                + relativedelta(months=1)
                - relativedelta(days=1)
            )
    return last_coupon


def get_coupon_schedule(maturity, freq, settlement):
    delta_months = 12 // freq
    coupon_dates = []
    next_coupon = maturity
    while next_coupon > settlement:
        coupon_dates.insert(0, next_coupon)
        next_coupon = add_months(next_coupon, -delta_months)
    return coupon_dates


def year_fraction(start_date, end_date, convention):
    days = (end_date - start_date).days

    if convention == "30/360":
        d1 = min(start_date.day, 30)
        d2 = min(end_date.day, 30) if d1 == 30 else end_date.day
        m1, m2 = start_date.month, end_date.month
        y1, y2 = start_date.year, end_date.year
        return ((360 * (y2 - y1)) + (30 * (m2 - m1)) + (d2 - d1)) / 360

    elif convention == "ACT/360":
        return days / 360

    elif convention == "ACT/365":
        return days / 365

    elif convention == "ACT/ACT":
        yf = 0.0
        current = start_date
        while current < end_date:
            year_end = min(datetime(current.year + 1, 1, 1), end_date)
            year_days = 366 if calendar.isleap(current.year) else 365
            yf += (year_end - current).days / year_days
            current = year_end
        return yf

    else:
        raise ValueError("Unsupported day count convention")


def year_fraction_ACT_ACT(start_date, end_date):
    current = start_date
    yf = 0.0
    while current < end_date:
        year_end = min(datetime(current.year + 1, 1, 1), end_date)
        year_days = 366 if calendar.isleap(current.year) else 365
        yf += (year_end - current).days / year_days
        current = year_end
    return yf


def calculate_accrued_interest(settlement, maturity, coupon_amount, freq, ex_days, convention):
    last_coupon = get_last_coupon_date(settlement, freq, maturity)
    next_coupon = add_months(last_coupon, 12 // freq)

    in_ex_interest = (next_coupon - settlement).days <= ex_days

    if in_ex_interest:
        last_coupon = next_coupon
        next_coupon = add_months(last_coupon, 12 // freq)

    if convention == "30/360":
        d1 = min(settlement.day, 30)
        d2 = min(last_coupon.day, 30)
        accrued_days = (
            (settlement.year - last_coupon.year) * 360
            + (settlement.month - last_coupon.month) * 30
            + (d1 - d2)
        )
        period_days = 360 // freq

    elif convention == "ACT/360":
        accrued_days = (settlement - last_coupon).days
        period_days = 360 // freq

    else:
        accrued_days = (settlement - last_coupon).days
        period_days = (next_coupon - last_coupon).days

    if period_days == 0:
        return Decimal("0.0")

    return Decimal(coupon_amount) * Decimal(accrued_days) / Decimal(period_days)


# =========================
# Core calculation
# =========================
def run_bond_calculation(
    settlement,
    maturity,
    ytm_pct,
    coupon_rate_pct,
    face_value,
    freq,
    interest_status,
    convention,
    ex_days,
    rounding_mode,
):
    if settlement >= maturity:
        raise ValueError("Settlement Date must be before Maturity Date.")

    ytm = float(ytm_pct) / 100
    coupon_rate = float(coupon_rate_pct) / 100
    face_value = float(face_value)

    is_ex_interest_selected = interest_status == "Ex"
    coupon_amount = (coupon_rate * face_value) / freq
    delta_months = 12 // freq

    payment_dates = []
    next_coupon = maturity
    while True:
        payment_dates.insert(0, next_coupon)
        next_coupon = add_months(next_coupon, -delta_months)
        if next_coupon <= settlement:
            break

    if payment_dates and payment_dates[0] <= settlement:
        payment_dates = payment_dates[1:]

    if not payment_dates:
        raise ValueError("No future payment dates found.")

    price = 0.0
    times, cashflows, discount_factors = [], [], []

    is_bill = len(payment_dates) == 1 and payment_dates[0] == maturity
    days_to_next = (payment_dates[0] - settlement).days

    if is_bill:
        if convention in ["30/360", "ACT/360"]:
            short_t = days_to_next / 360
        elif convention == "ACT/365":
            short_t = days_to_next / 365
        else:
            short_t = year_fraction_ACT_ACT(settlement, payment_dates[0])
    else:
        anchor_day = payment_dates[0].day
        last_coupon = add_months(payment_dates[0], -delta_months)
        try:
            last_coupon = last_coupon.replace(day=anchor_day)
        except ValueError:
            last_coupon = (
                last_coupon.replace(day=1)
                + relativedelta(months=1)
                - relativedelta(days=1)
            )

        full_coupon_period_days = (payment_dates[0] - last_coupon).days
        if full_coupon_period_days == 0:
            raise ValueError("Coupon period length is zero.")

        short_t = days_to_next / full_coupon_period_days

    if is_bill:
        short_df = 1 / (1 + ytm * short_t)
    else:
        short_df = (1 / (1 + ytm / freq)) ** short_t

    amort_rows = []

    for i, dt in enumerate(payment_dates):
        n = i

        if is_bill:
            cf = coupon_amount + face_value
            df = short_df
            pv = cf * df
            label = "*** FINAL PAYMENT ***"
            times.append(short_t)
        else:
            df_long = 1 / (1 + ytm / freq) ** n
            df = df_long * short_df

            cf = coupon_amount
            if dt == maturity:
                cf += face_value

            pv = cf * df
            label = "*** FINAL PAYMENT ***" if dt == maturity else ""

        if i == 0 and is_ex_interest_selected and not is_bill:
            label += " (EXCLUDED FROM PRICE)"
        else:
            price += pv

        times.append(n / freq)
        cashflows.append(cf)
        discount_factors.append(df)

        amort_rows.append(
            {
                "Date": dt.strftime("%d/%m/%Y"),
                "Cashflow": round(cf, 2),
                "Discount Factor": round(df, 10),
                "Present Value": round(pv, 2),
                "Note": label,
            }
        )

    if is_ex_interest_selected:
        accrued_interest = (
            calculate_accrued_interest(settlement, maturity, coupon_amount, freq, ex_days, convention)
            - Decimal(coupon_amount)
        )
    else:
        accrued_interest = calculate_accrued_interest(
            settlement, maturity, coupon_amount, freq, ex_days, convention
        )

    gross_price_per_100 = Decimal(price) / Decimal(face_value) * Decimal(100)
    accrued_per_100 = Decimal(accrued_interest) / Decimal(face_value) * Decimal(100)

    if rounding_mode == "quantum":
        if is_bill:
            price_for_settlement = round(gross_price_per_100, 12)
        else:
            price_for_settlement = round(gross_price_per_100, 3)
        accrued_per_100_rounded = round(accrued_per_100, 3)

    elif rounding_mode == "12dp":
        price_for_settlement = gross_price_per_100
        accrued_per_100_rounded = accrued_per_100

    else:
        price_for_settlement = round(gross_price_per_100, 3)
        accrued_per_100_rounded = round(accrued_per_100, 3)

    total_settlement = price_for_settlement * Decimal(face_value) / Decimal(100)
    accrued_interest_amount = accrued_per_100_rounded * Decimal(face_value) / Decimal(100)
    capital_amount = total_settlement - accrued_interest_amount

    clean_display = float(round(gross_price_per_100 - accrued_per_100_rounded, 3))
    accrued_display = float(round(accrued_per_100_rounded, 3))
    gross_display = float(round(gross_price_per_100, 3))

    risk_cashflows = []
    risk_times = []
    risk_discount_factors = []

    for dt in payment_dates:
        t = year_fraction(settlement, dt, convention) if not is_bill else short_t

        cf = coupon_amount if not is_bill else coupon_amount + face_value
        if dt == maturity:
            cf += face_value if not is_bill else 0

        df = (1 / (1 + ytm / freq) ** (t * freq)) if not is_bill else short_df

        risk_cashflows.append(cf)
        risk_times.append(t)
        risk_discount_factors.append(df)

    dirty_price = sum(cf * df for cf, df in zip(risk_cashflows, risk_discount_factors))
    macaulay_duration = compute_macaulay_duration(
        risk_cashflows, risk_discount_factors, risk_times, dirty_price
    )
    modified_duration = macaulay_duration / (1 + ytm / freq)
    convexity = compute_convexity(
        risk_cashflows, risk_times, ytm, dirty_price, freq, is_bill
    )

    in_ex_interest_period = (payment_dates[0] - settlement).days <= ex_days
    info_message = ""
    if is_bill:
        info_message = "PRICED AS A BILL"
    elif in_ex_interest_period:
        info_message = "(Ex-Interest on old issues / set new issues to cum interest)"

    return {
        "clean_price_per_100": clean_display,
        "accrued_per_100": accrued_display,
        "gross_price_per_100": gross_display,
        "capital_amount": float(capital_amount),
        "accrued_interest_amount": float(accrued_interest_amount),
        "total_settlement": float(total_settlement),
        "duration": float(macaulay_duration),
        "modified_duration": float(modified_duration),
        "convexity": float(convexity),
        "amort_df": pd.DataFrame(amort_rows),
        "info_message": info_message,
        "is_bill": is_bill,
        "coupon_amount": coupon_amount,
    }


def generate_chart(result, settlement, maturity, ytm_pct, coupon_rate_pct, face_value, freq, convention, ex_days, interest_status):
    ytm = ytm_pct / 100
    coupon_rate = coupon_rate_pct / 100
    coupon_amt = (coupon_rate * face_value) / freq
    coupon_dates = get_coupon_schedule(maturity, freq, settlement)
    is_bill = len(coupon_dates) == 1 and coupon_dates[0] == maturity
    is_ex_interest_selected = interest_status == "Ex"

    def bond_pv(y):
        if is_bill:
            t = year_fraction(settlement, maturity, convention)
            df = 1 / (1 + y * t)
            pv = (face_value + coupon_amt) * df
            macaulay = t
            modified = macaulay / (1 + y)
            convexity = (t * (t + 1)) / ((1 + y) ** 2)
            return pv, macaulay, modified, convexity

        pv = 0.0
        dirty_pv = 0.0
        risk_cf, risk_t, risk_df = [], [], []

        for i, dt in enumerate(coupon_dates):
            if dt <= settlement:
                continue

            t = year_fraction(settlement, dt, convention)
            cf = coupon_amt
            if dt == maturity:
                cf += face_value

            df = 1 / (1 + y / freq) ** (t * freq)
            risk_cf.append(cf)
            risk_t.append(t)
            risk_df.append(df)
            dirty_pv += cf * df

            if i == 0 and is_ex_interest_selected:
                continue

            pv += cf * df

        macaulay = compute_macaulay_duration(risk_cf, risk_df, risk_t, dirty_pv) if dirty_pv else 0
        modified = macaulay / (1 + y / freq) if dirty_pv else 0
        convexity = compute_convexity(risk_cf, risk_t, y, dirty_pv, freq, False) if dirty_pv else 0

        return pv, macaulay, modified, convexity

    min_yield = max(0.0001, ytm - 0.025)
    max_yield = ytm + 0.025
    yields = np.linspace(min_yield, max_yield, 100)

    prices = []
    convexities = []
    for y in yields:
        dirty, _, _, cx = bond_pv(y)
        prices.append(dirty)
        convexities.append(cx)

    x_vals = np.array([y * 100 for y in yields])

    fig, ax = plt.subplots(figsize=(10, 5))
    ax2 = ax.twinx()

    ax.plot(x_vals, prices, label="Dirty Price Curve")
    ax.axhline(result["total_settlement"], linestyle="--", label=f'Observed Dirty Price: ${result["total_settlement"]:,.2f}')
    ax.axvline(ytm * 100, linestyle="--", label=f"Yield: {ytm * 100:.2f}%")
    ax.scatter(ytm * 100, result["total_settlement"], label="Current Price Point")

    ax2.plot(x_vals, convexities, label="Convexity Curve")

    ax.set_title("Bond Dirty Price and Convexity vs Yield to Maturity")
    ax.set_xlabel("Yield to Maturity (%)")
    ax.set_ylabel("Dirty Price")
    ax2.set_ylabel("Convexity")
    ax.grid(True)

    ax.legend(loc="upper left")
    ax2.legend(loc="upper right")

    return fig


class PDFWithFooter(FPDF):
    def footer(self):
        self.set_y(-12)
        self.set_font("Arial", size=8)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")


def build_pdf_bytes(inputs, result):
    pdf = PDFWithFooter()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Northern Territory Treasury Corporation", ln=True, align="C")
    pdf.cell(0, 10, "Bond Price Calculator Results", ln=True, align="C")

    pdf.ln(5)
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 10, f"Printed on: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ln=True)

    pdf.ln(2)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "--- Input Data ---", ln=True)
    pdf.set_font("Arial", size=10)

    for k, v in inputs.items():
        pdf.cell(0, 8, f"{k}: {v}", ln=True)

    pdf.ln(2)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "--- Price per $100 ---", ln=True)
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 8, f'Capital Price: {result["clean_price_per_100"]:.3f}', ln=True)
    pdf.cell(0, 8, f'Accrued Interest: {result["accrued_per_100"]:.3f}', ln=True)
    pdf.cell(0, 8, f'Gross Price: {result["gross_price_per_100"]:.3f}', ln=True)

    pdf.ln(2)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "--- Settlement Values ---", ln=True)
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 8, f'Capital Amount: ${result["capital_amount"]:,.2f}', ln=True)
    pdf.cell(0, 8, f'Accrued Interest Amount: ${result["accrued_interest_amount"]:,.2f}', ln=True)
    pdf.cell(0, 8, f'Total Settlement Value: ${result["total_settlement"]:,.2f}', ln=True)

    pdf.ln(2)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "--- Risk Metrics ---", ln=True)
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 8, f'Duration: {result["duration"]:.3f}', ln=True)
    pdf.cell(0, 8, f'Modified Duration: {result["modified_duration"]:.3f}', ln=True)
    pdf.cell(0, 8, f'Convexity: {result["convexity"]:.3f}', ln=True)

    pdf.ln(2)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "--- Amortization Schedule ---", ln=True)
    pdf.set_font("Courier", size=8)

    for _, row in result["amort_df"].iterrows():
        line = f'{row["Date"]:15} {row["Cashflow"]:15,.2f} {row["Discount Factor"]:18,.10f} {row["Present Value"]:18,.2f} {row["Note"]}'
        pdf.multi_cell(0, 5, line)

    pdf_bytes = pdf.output(dest="S").encode("latin-1")
    return pdf_bytes


# =========================
# UI
# =========================
st.title("BOND PRICE CALCULATOR")

col1, col2 = st.columns(2)

with col1:
    settlement_date = st.date_input("Settlement Date", value=date.today())
    maturity_date = st.date_input("Maturity Date", value=date.today() + relativedelta(years=5))
    ytm_pct = st.number_input("Yield (% per annum)", value=5.000000, format="%.6f")
    coupon_rate_pct = st.number_input("Coupon Rate (% of face value)", value=5.000000, format="%.6f")
    face_value = st.number_input("Face Value ($)", min_value=0.0, value=100000.0, step=1000.0, format="%.2f")

with col2:
    freq = int(st.selectbox("Coupon Frequency", [1, 2, 4], index=0))
    interest_status = st.selectbox("Interest Status", ["Cum", "Ex"], index=0)
    convention = st.selectbox("Day Count Convention", ["ACT/ACT", "30/360", "ACT/360", "ACT/365"], index=0)
    ex_days = int(st.number_input("Ex-Interest Period (days)", min_value=0, max_value=30, value=7))
    rounding_mode = st.radio("Rounding", ["12dp", "3dp", "quantum"], index=1, horizontal=True)

calc_col, chart_col, pdf_col = st.columns(3)
calculate_clicked = calc_col.button("Calculate", use_container_width=True)
chart_clicked = chart_col.button("Generate Chart", use_container_width=True)
pdf_clicked = pdf_col.button("Prepare PDF", use_container_width=True)

if calculate_clicked or chart_clicked or pdf_clicked:
    try:
        settlement = datetime.combine(settlement_date, datetime.min.time())
        maturity = datetime.combine(maturity_date, datetime.min.time())

        result = run_bond_calculation(
            settlement=settlement,
            maturity=maturity,
            ytm_pct=ytm_pct,
            coupon_rate_pct=coupon_rate_pct,
            face_value=face_value,
            freq=freq,
            interest_status=interest_status,
            convention=convention,
            ex_days=ex_days,
            rounding_mode=rounding_mode,
        )

        if result["info_message"]:
            st.warning(result["info_message"])

        c1, c2, c3 = st.columns(3)

        with c1:
            st.markdown("### Price per $100")
            st.metric("Capital Price", f'{result["clean_price_per_100"]:.3f}')
            st.metric("Accrued Interest", f'{result["accrued_per_100"]:.3f}')
            st.metric("Gross Price", f'{result["gross_price_per_100"]:.3f}')

        with c2:
            st.markdown("### Settlement")
            st.metric("Capital Amount", f'${result["capital_amount"]:,.2f}')
            st.metric("Accrued Interest Amount", f'${result["accrued_interest_amount"]:,.2f}')
            st.metric("Total Settlement Value", f'${result["total_settlement"]:,.2f}')

        with c3:
            st.markdown("### Risk Metrics")
            st.metric("Duration", f'{result["duration"]:.3f}')
            st.metric("Modified Duration", f'{result["modified_duration"]:.3f}')
            st.metric("Convexity", f'{result["convexity"]:.3f}')

        st.markdown("### Amortization Schedule")
        st.dataframe(result["amort_df"], use_container_width=True)

        if chart_clicked:
            fig = generate_chart(
                result=result,
                settlement=settlement,
                maturity=maturity,
                ytm_pct=ytm_pct,
                coupon_rate_pct=coupon_rate_pct,
                face_value=face_value,
                freq=freq,
                convention=convention,
                ex_days=ex_days,
                interest_status=interest_status,
            )
            st.pyplot(fig)

        if pdf_clicked:
            inputs = {
                "Settlement Date": settlement.strftime("%d/%m/%Y"),
                "Maturity Date": maturity.strftime("%d/%m/%Y"),
                "Yield": f"{ytm_pct:.6f}%",
                "Coupon Rate": f"{coupon_rate_pct:.6f}%",
                "Face Value": f"${face_value:,.2f}",
                "Coupon Frequency": str(freq),
                "Interest Status": interest_status,
                "Day Count Convention": convention,
                "Ex-Interest Period (days)": str(ex_days),
                "Rounding": rounding_mode,
            }

            pdf_bytes = build_pdf_bytes(inputs, result)
            st.download_button(
                label="Download PDF",
                data=pdf_bytes,
                file_name="bond_calculator_results.pdf",
                mime="application/pdf",
            )

    except Exception as e:
        st.error(f"Calculation Error: {str(e)}")
