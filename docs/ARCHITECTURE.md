# Architecture Overview

## 🇬🇧 English

### General Flow

User Selection  
↓  
Lock Mechanism  
↓  
Execution Path  

**AD On-Prem**
- Account Disable
- Group Removal
- Organization Cleanup (Manager + Direct Reports only)

**Exchange Online**
- Convert to Shared Mailbox
- Calendar Cleanup (3 retry logic)

**Purview**
- Create Content Search
- Start Search

**Microsoft Graph**
- Revoke Sessions
- Remove Group Memberships
- Remove Licenses

---

### Security Controls

Execution is automatically blocked if the selected user is a member of:

- Domain Admins
- Enterprise Admins
- Schema Admins
- Built-in Administrators

This prevents accidental modification of privileged accounts.

---

### Execution Guard

Actions remain disabled until:
- A user is selected
- Lock is activated

This prevents accidental execution.



---

## 🇹🇷 Türkçe

### Genel Akış

Kullanıcı Seçimi  
↓  
Kilitleme Mekanizması  
↓  
İşlem Yürütme  

**AD On-Prem**
- Hesabı devre dışı bırakma
- Gruplardan çıkarma
- Organizasyon temizliği (Sadece Manager + Direct Reports)

**Exchange Online**
- Shared Mailbox'a çevirme
- Takvim temizliği (3 deneme mekanizması)

**Purview**
- Content Search oluşturma
- Aramayı başlatma

**Microsoft Graph**
- Oturumları düşürme
- M365 gruplarından çıkarma
- Lisans kaldırma

---

### Güvenlik Kontrolleri

Seçilen kullanıcı aşağıdaki gruplardan birine üyeyse işlem otomatik olarak engellenir:

- Domain Admins
- Enterprise Admins
- Schema Admins
- Built-in Administrators

Bu mekanizma kritik hesapların yanlışlıkla etkilenmesini önler.

---

### İşlem Koruması

Aşağıdaki şartlar sağlanmadan işlem butonları aktif olmaz:

- Kullanıcı seçilmelidir
- Kilitle butonu aktif edilmelidir

Bu sayede yanlışlıkla işlem yapılması engellenir.
