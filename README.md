

# Offboarding Automation (PowerShell WinForms) - v1

## 🇬🇧 English

### Overview
Offboarding Automation is a PowerShell WinForms-based tool that centralizes common offboarding operations in a single interface.

It supports:

- Active Directory (On-Prem) actions  
- Exchange Online actions  
- Purview Content Search (Search only, export via UI)  
- Microsoft Graph / Entra ID actions  
- TR / EN language switch  
- Privileged account protection (Domain/Enterprise/Schema Admins + Built-in Administrators block)

---

### Features

- Disable AD account
- Remove from AD groups
- Reset password
- Update description
- Hide from GAL
- Organization cleanup (Manager + Direct Reports only)
- Convert mailbox to shared
- Remove calendar events (3 retry logic)
- Create and start Purview Content Search
- Revoke sign-in sessions
- Remove M365 group memberships
- Remove licenses
- Secure lock mechanism before execution

---

### Requirements

- Windows PowerShell 5.1+
- RSAT / ActiveDirectory module
- ExchangeOnlineManagement module
- Microsoft.Graph module
- Proper permissions for:
  - Exchange Online
  - Microsoft Graph
  - Purview Content Search

---

### Quick Start

#### Option A — Run as Administrator (Recommended)
1. Open the `src` folder
2. Run `run_offboarding_admin.bat`

#### Option B — Run via PowerShell

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
.\src\OffboardingAutomation.ps1
```
--

# 🇹🇷 Türkçe

## Genel Bakış

Offboarding Automation, kullanıcı çıkış süreçlerini tek bir arayüzde toplayan PowerShell WinForms tabanlı bir otomasyon aracıdır.

Desteklenen işlemler:

- Active Directory (On-Prem) işlemleri  
- Exchange Online işlemleri  
- Purview Content Search (Sadece arama, export arayüz üzerinden)  
- Microsoft Graph / Entra ID işlemleri  
- TR / EN dil seçimi  
- Kritik admin hesapları için güvenlik engeli  

---

## Özellikler

- AD hesabını devre dışı bırakma  
- AD gruplarından çıkarma  
- Parola sıfırlama  
- Description güncelleme  
- GAL’den gizleme  
- Organizasyon temizliği (Sadece Manager + Direct Reports)  
- Mailbox’ı Shared’a çevirme  
- Takvim etkinliklerini silme (3 deneme mekanizması)  
- Purview arama oluşturma ve başlatma  
- Oturumları düşürme (Revoke)  
- M365 gruplarından çıkarma  
- Lisans geri alma  
- İşlem öncesi kilitleme güvenliği  

---

## Gereksinimler

- Windows PowerShell 5.1+  
- RSAT / ActiveDirectory modülü  
- ExchangeOnlineManagement modülü  
- Microsoft.Graph modülü  
- Exchange, Graph ve Purview için gerekli yetkiler  

---

## Hızlı Başlangıç

### Seçenek A — Yönetici Olarak Çalıştır (Önerilen)

1. `src` klasörüne girin  
2. `run_offboarding_admin.bat` dosyasını çalıştırın  

### Seçenek B — PowerShell ile Çalıştırma

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
.\src\OffboardingAutomation.ps1
```

Nasıl Kullanılır:

- En az 2 karakter ile kullanıcı arayın

- Listeden doğru kullanıcıyı seçin

- Kilitle (Lock) butonuna basın

İstediğiniz işlemi seçin:

- SADECE AD

- SADECE EXO/PURVIEW

- SADECE GRAPH

- HEPSİNİ ÇALIŞTIR

- Log panelinden süreci takip edin

- Purview PST Export

Bu araç Purview üzerinde Content Search oluşturur ve başlatır.

PST export almak için:

https://purview.microsoft.com/
 adresine giriş yapın

Dava Dosyaları > İçerik Arama bölümüne gidin

Uygulama logunda görünen arama adını açın

Dışarı Aktar seçeneği ile PST paketini oluşturun

Not: PowerShell ile export desteği Microsoft tarafından kaldırılmıştır. Export işlemi arayüz üzerinden yapılır.

Lisans

MIT
