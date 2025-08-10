import streamlit as st
from bs4 import BeautifulSoup
import pandas as pd
import re
from datetime import datetime, timedelta

def convert_utc_to_brt(utc_str):
    try:
        dt = datetime.fromisoformat(utc_str.replace("Z", "+00:00"))
        return (dt - timedelta(hours=3)).strftime("%Y-%m-%d %H:%M:%S")
    except:
        return None

def extract_data(html_content):
    soup = BeautifulSoup(html_content, "html.parser")
    text = soup.get_text(" ")
    phones = re.findall(r'\+?\d{1,3}[\s\-]?\(?\d{2,3}\)?[\s\-]?\d{4,5}[\s\-]?\d{4}', text)
    ipv4_matches = re.findall(r'((?:\d{1,3}\.){3}\d{1,3}).{0,40}?(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z)', text)
    ipv6_matches = re.findall(r'((?:[A-F0-9]{1,4}:){7}[A-F0-9]{1,4}).{0,40}?(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z)', text, re.IGNORECASE)
    ipv4 = [(ip, ts, convert_utc_to_brt(ts)) for ip, ts in ipv4_matches]
    ipv6 = [(ip, ts, convert_utc_to_brt(ts)) for ip, ts in ipv6_matches]
    return phones, ipv4, ipv6

st.title("Relatório Policial Automático")
files = st.file_uploader("Carregar arquivos HTML", type="html", accept_multiple_files=True)

if files:
    all_phones, all_ipv4, all_ipv6 = [], [], []
    for f in files:
        content = f.read().decode("utf-8")
        phones, ipv4, ipv6 = extract_data(content)
        all_phones.extend(phones)
        all_ipv4.extend(ipv4)
        all_ipv6.extend(ipv6)

    st.subheader("Telefones")
    st.dataframe(pd.DataFrame(list(set(all_phones)), columns=["Telefone"]))

    st.subheader("IPv4 e horários")
    st.dataframe(pd.DataFrame(all_ipv4, columns=["IPv4", "Data/Hora UTC", "Data/Hora UTC-3"]))

    st.subheader("IPv6 e horários")
    st.dataframe(pd.DataFrame(all_ipv6, columns=["IPv6", "Data/Hora UTC", "Data/Hora UTC-3"]))
