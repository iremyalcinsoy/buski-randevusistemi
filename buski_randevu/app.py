import io
import os
import pandas as pd
import openpyxl
import pyodbc  # EKLENDİ: SQL Server için
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file
from flask_mail import Mail, Message

app = Flask(__name__)
app.secret_key = "buski_secret_key_12345"

# GMAIL SMTP AYARLARI
SYSTEM_EMAIL = 'iremyalcinsoyyy669@gmail.com'

app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USE_SSL'] = False
app.config['MAIL_USERNAME'] = SYSTEM_EMAIL 
app.config['MAIL_PASSWORD'] = 'mzogolwzbgltoadh'  # Google Uygulama Şifresi
app.config['MAIL_DEFAULT_SENDER'] = ('BUSKİ Randevu Sistemi', SYSTEM_EMAIL)

mail = Mail(app)

# ==========================================
# SQL SERVER BAĞLANTI AYARLARI (EKLENDİ)
# ==========================================
def get_db_connection():
    conn_str = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=.\\SQLEXPRESS01;"
        "DATABASE=BuskiRandevuDB;"
        "Trusted_Connection=yes;"
    )
    return pyodbc.connect(conn_str)

# Veritabanı satırlarını sözlüğe (dictionary) çeviren yardımcı fonksiyon (HTML bozulmasın diye)
def rows_to_dict_list(cursor):
    columns = [column[0] for column in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


# 1. ANA SAYFA ROTASI
@app.route('/')
def index():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Randevular")
    randevular = rows_to_dict_list(cursor)
    conn.close()
    
    return render_template('index.html', randevular=randevular)


# 2. RANDEVU OLUŞTURMA ROTASI 
@app.route('/randevu-al', methods=['POST'])
def randevu_al():
    islem_turu = request.form.get('islem_turu')
    tc_no = request.form.get('tc_no')
    ad_soyad = request.form.get('ad_soyad')
    user_email = request.form.get('email')
    tarih = request.form.get('tarih')
    saat = request.form.get('saat')

    conn = get_db_connection()
    cursor = conn.cursor()

    # DOLULUK KONTROLU (Veritabanından sorgulama)
    cursor.execute("""
        SELECT * FROM Randevular 
        WHERE tarih=? AND saat=? AND islem_turu=? AND durum != 'İptal Edildi'
    """, (tarih, saat, islem_turu))
    
    ayni_zamanli_randevu = cursor.fetchone()

    if ayni_zamanli_randevu:
        conn.close()
        flash(f"Üzgünüz, seçtiğiniz {tarih} tarihi, {saat} saati ve '{islem_turu}' için kontenjan doludur. Lütfen başka bir gün, saat veya işlem türü seçiniz.", "danger")
        return redirect(url_for('index'))
    # --------------------------------------------------------------------------

    # Yeni randevuyu veritabanına ekle
    cursor.execute("""
        INSERT INTO Randevular (tc_no, ad_soyad, email, islem_turu, tarih, saat, durum)
        VALUES (?, ?, ?, ?, ?, ?, 'Beklemede')
    """, (tc_no, ad_soyad, user_email, islem_turu, tarih, saat))
    conn.commit()
    conn.close()

    #KULLANICIYA ONAY MAILI GONDERME
    try:
        msg = Message(
            subject=f"BUSKİ Randevu Onayı - Sayın {ad_soyad}",
            recipients=[user_email]
        )
        msg.body = f"""Sayın {ad_soyad},

BUSKİ Online Randevu Sisteminden randevunuz başarıyla oluşturulmuştur.

Randevu Detayları:
- İşlem Türü: {islem_turu}
- Randevu Tarihi: {tarih}
- Randevu Saati: {saat}
- T.C. Kimlik No: {tc_no}

Randevunuzu sorgulamak veya iptal etmek için web sitemizdeki "Randevu Sorgula / İptal" sayfasını kullanabilirsiniz.

Saygılarımızla,
BUSKİ Genel Müdürlüğü
"""
        mail.send(msg)
        flash(f"Randevunuz başarıyla oluşturuldu! Onay e-postası '{user_email}' adresine gönderildi.", "success")
    except Exception as e:
        print("Mail Gönderme Hatası:", e)
        flash(f"Randevu oluşturuldu ancak e-posta gönderilemedi. (Hata: {e})", "warning")

    return redirect(url_for('index'))


# ==========================================
# 3. RANDEVU SORGULA / İPTAL EKRANI ROTASI
# ==========================================
@app.route('/sorgula', methods=['GET', 'POST'])
def sorgula():
    bulunan_randevular = None
    if request.method == 'POST':
        tc_no = request.form.get('tc_no')
        
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Randevular WHERE tc_no = ?", (tc_no,))
        bulunan_randevular = rows_to_dict_list(cursor)
        conn.close()

        if not bulunan_randevular:
            flash("Bu T.C. Kimlik numarasına ait bir randevu bulunamadı.", "danger")

    return render_template('sorgula.html', randevular=bulunan_randevular)


# ==========================================
# 4. KULLANICI RANDEVU İPTAL ROTASI
# ==========================================
@app.route('/randevu-iptal/<int:randevu_id>', methods=['POST'])
def randevu_iptal(randevu_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # İptal edilecek randevunun bilgilerini al (Mail göndermek için)
    cursor.execute("SELECT * FROM Randevular WHERE id = ?", (randevu_id,))
    columns = [column[0] for column in cursor.description]
    row = cursor.fetchone()
    
    if row:
        randevu = dict(zip(columns, row))
        
        # Veritabanında durumu güncelle
        cursor.execute("UPDATE Randevular SET durum = 'İptal Edildi' WHERE id = ?", (randevu_id,))
        conn.commit()
        
        try:
            msg = Message(
                subject=f"BUSKİ - Randevu İptal Bilgilendirmesi",
                recipients=[randevu['email']]
            )
            msg.body = f"""Sayın {randevu['ad_soyad']},

{randevu['tarih']} tarihindeki {randevu['islem_turu']} randevunuz talebiniz üzerine İPTAL EDİLMİŞTİR.

BUSKİ Genel Müdürlüğü
"""
            mail.send(msg)
            flash("Randevunuz iptal edildi ve bilgilendirme e-postası gönderildi.", "success")
        except Exception as e:
            flash("Randevu iptal edildi fakat e-posta gönderilemedi.", "warning")
    else:
        flash("İptal edilecek randevu bulunamadı.", "danger")

    conn.close()
    return redirect(url_for('sorgula'))


# ==========================================
# 5. ADMIN GİRİŞİ VE DASHBOARD ROTALARI
# ==========================================
@app.route('/admin-login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        if username == "admin" and password == "123456":
            flash("Admin paneline başarıyla giriş yapıldı.", "success")
            return redirect(url_for('admin_dashboard'))
        else:
            flash("Hatalı Kullanıcı Adı veya Şifre!", "danger")

    return render_template('admin_login.html')


@app.route('/admin-dashboard')
def admin_dashboard():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Randevular")
    randevular = rows_to_dict_list(cursor)
    conn.close()

    toplam_randevu = len(randevular)
    bekleyen = sum(1 for r in randevular if r['durum'] == 'Beklemede' or r['durum'] == 'Bekliyor')
    gelen = sum(1 for r in randevular if r['durum'] == 'Geldi')
    iptal = sum(1 for r in randevular if r['durum'] == 'İptal Edildi')

    return render_template(
        'admin_dashboard.html', 
        randevular=randevular,
        toplam_randevu=toplam_randevu,
        bekleyen=bekleyen,
        gelen=gelen,
        iptal=iptal
    )


@app.route('/admin-logout')
def admin_logout():
    flash("Oturum kapatıldı.", "info")
    return redirect(url_for('admin_login'))


# ==========================================
# 6. ADMIN - RANDEVU İPTAL ROTASI
# ==========================================
@app.route('/admin-randevu-iptal/<int:randevu_id>', methods=['POST'])
def admin_randevu_iptal(randevu_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT * FROM Randevular WHERE id = ?", (randevu_id,))
    columns = [column[0] for column in cursor.description]
    row = cursor.fetchone()
    
    if row:
        randevu = dict(zip(columns, row))
        yeni_durum = request.form.get('durum')
        
        if yeni_durum:
            cursor.execute("UPDATE Randevular SET durum = ? WHERE id = ?", (yeni_durum, randevu_id))
            conn.commit()
            randevu['durum'] = yeni_durum # Mail taslağı için güncelliyoruz
        
        try:
            msg = Message(
                subject=f"BUSKİ - Randevu Güncelleme Bilgilendirmesi (Yönetici Tarafından)",
                recipients=[randevu['email']]
            )
            msg.body = f"""Sayın {randevu['ad_soyad']},

{randevu['tarih']} tarihindeki {randevu['islem_turu']} randevunuzun durumu kurumumuz yöneticisi tarafından '{randevu['durum']}' olarak güncellenmiştir.

Detaylı bilgi için kurumumuzla iletişime geçebilirsiniz.

BUSKİ Genel Müdürlüğü
"""
            mail.send(msg)
            flash(f"#{randevu_id} numaralı randevu durumu güncellendi ve kullanıcıya mail atıldı.", "success")
        except Exception as e:
            flash(f"Randevu güncellendi fakat e-posta gönderilirken bir hata oluştu: {e}", "warning")
    else:
        flash("Güncellenecek randevu bulunamadı.", "danger")

    conn.close()
    return redirect(url_for('admin_dashboard'))



# 7. ADMIN - EXCEL RAPORU INDIRME ROTASI (LOGOLU VE BASLIKLI)
@app.route('/admin/excel-indir')
def excel_indir():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Randevular")
    randevular = rows_to_dict_list(cursor)
    conn.close()

    # Sadece Excel basliklarinin duzgun gorunmesi icin anahtar eslemesi yapildi
    data_list = []
    for r in randevular:
        data_list.append({
            "id": r["id"],
            "tc no": r["tc_no"],
            "ad / soyad": r["ad_soyad"],
            "email": r["email"],
            "işlem türü": r["islem_turu"],
            "tarih": r["tarih"],
            "saat": r["saat"],
            "durum": r["durum"]
        })

    df = pd.DataFrame(data_list)
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Randevular', startrow=5)
        
        workbook = writer.book
        worksheet = writer.sheets['Randevular']
        
        # 1. satirdan 5. satira kadar bosluklari temizle ve cizgileri (border) tamamen kaldir
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        
        for row in range(1, 6):
            for col in range(1, 20):  # Sag taraftaki tum olasi hucre sinirlarini da kapsayacak sekilde genisletildi
                cell = worksheet.cell(row=row, column=col)
                cell.fill = white_fill
                cell.border = Border()

        # Logo B2 hucresine yerlestirildi
        logo_path = os.path.join('static', 'logo.png')
        if os.path.exists(logo_path):
            img = Image(logo_path)
            img.width = 110
            img.height = 50
            worksheet.add_image(img, 'B2')

        # Ust Kurum Basliklari
        worksheet['C2'] = "BURSA BÜYÜKŞEHİR BELEDİYESİ - BUSKİ GENEL MÜDÜRLÜĞÜ"
        worksheet['C2'].font = Font(name="Calibri", size=12, bold=True, color="0B1329")
        
        worksheet['C3'] = "Online Randevu Yönetim Sistemi - Rapor Listesi"
        worksheet['C3'].font = Font(name="Calibri", size=10, italic=True, color="555555")

        # Tablo Sutun Basliklari Tasarimi (6. Satir)
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_alignment = Alignment(horizontal="center", vertical="center")
        thin_border = Border(
            left=Side(style='thin', color='CCCCCC'),
            right=Side(style='thin', color='CCCCCC'),
            top=Side(style='thin', color='CCCCCC'),
            bottom=Side(style='thin', color='CCCCCC')
        )

        for col_num in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=6, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = thin_border

        # Sütun Genişlikleri (Sağ taraftaki boşluk çizgilerinin görünmemesi için genişletildi)
        custom_widths = {
            'A': 12,  # id
            'B': 20,  # tc_no
            'C': 30,  # ad_soyad
            'D': 40,  # email
            'E': 35,  # islem_turu
            'F': 18,  # tarih
            'G': 14,  # saat
            'H': 22,  # durum
            'I': 20,  # Fazladan sağ sütunlar (çizgilerin kaybolması için)
            'J': 20
        }
        
        for col_letter, width in custom_widths.items():
            worksheet.column_dimensions[col_letter].width = width

    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='buski_randevular.xlsx'
    )


# ==========================================
# SUNUCUYU BAŞLATMA
# ==========================================
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)