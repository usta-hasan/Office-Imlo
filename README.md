# Office Imlo — 2.5.0

2.5.0: Visio, Project va klassik Outlook uçun COM qöşimça qöşildi.
Visio: tanlangan matn/şakllar yoki barça sahifalar (jonli maydonlar saqlanadi).
Project: tanlangan Name/Text1–Text30 kataklari; formula/sana/raqamlar özgarmaydi.
Outlook: alohida xat yoziş oynasida tanlangan matn yoki xat tanasi; yuboriş yoq.
Barçasida usta_hasan > imlo: Alt+1 belgi, Alt+2 yangi, Alt+3 eski,
Alt+4 lotin→kiril, Alt+5 kiril→lotin; Alt+F1 ma’lumot.
Tarqatiş: Install.exe yoki OfficeImlo_2.5.0_Release.zip.
Developer Backup ZIP manba kodni saqlaydi — ommaviy tarqatmang.

Word, Excel va PowerPoint: Alt+1 — belgi tuzatiş; Alt+2 — Yangi alfavit; Alt+4 — lotindan kirilga; Alt+5 — kirildan lotinga. Word: Alt+6 — kitobça hisoblagiç. Barça Office dasturlarida Alt+F1 — ma’lumot. Raqamlar klaviaturaning yuqori qatoridan bosiladi.


## Örnatiş
Word, Excel va PowerPoint oynalarini yoping, Install_KochishBelgisi.exe faylini oçing va Örnatiş tugmasini bosing.
Ular qayta oçilganda `usta_hasan > imlo` bölimida Köçiş belgisi, Yangi alfavit, Lotin → Kiril va Kiril → Lotin tugmalari böladi. Kitobça hisoblagiç faqat Word'da qoladi.

Excel'da tanlangan kataklar (formulalar özgarmaydi), PowerPoint'da tanlangan matn/şakllar yoki tanlov bölmasa barça slaydlar qayta işlanadi.
- Köçiş belgisi
- Yangi alfavit: Ö/ö → Ö/ö; Ğ/ğ → Ğ/ğ; Ş/ş → Ş/ş; Ç/ç → Ç/ç.
- Lotin → Kiril
- Kiril → Lotin
- print → Kitobça hisoblagiç

Boşqa kompyuterga faqat örnatuvçi EXE faylini olib ötiş kifoya.
Windows, Word 2010+ va .NET Framework 4.x talab qilinadi. Tekşirilgan muhit: Word 16.0.

## 1.6.10 — SpinEdit vertikal markazi
`Bir tomondagi betlar` label'ining matni ham markazga tekislandi. Label, SpinEdit va `2 4 6 8 16` qatori endi aynan bitta vertikal markaziy çiziqda turadi.

## 1.6.9 — natija maydoni va SpinEdit tekisligi
Old/orqa natija maydonlarining scrollbar'i öçirildi. `Bir tomondagi betlar` label'i, SpinEdit va `2 4 6 8 16` qiymatlari bitta markaziy öq böyiça ustma-ust tekislandi. Boşqa forma elementlari özgartirilmadi.

## 1.6.8 — forma tekislaş va copy rasmi
`Hisob` GroupBox'i balandlaştirildi, `Yakuniy sahifa` label'i `Bölişi kerak` deb nomlandi. SpinEdit yuqoriga, 2/4/6/8/16 qiymatlari uning tagiga joylaştirildi. Pastki tugmalar kötarildi. Copy tugmalariga foydalanuvçi bergan rasm asosidagi bitmap töliq joylaştirildi; installer bitmapni `.dotm` bilan birga örnatadi va olib taşlaydi. Forma caption'i `usta_hasan - Print Calculation` böldi.

## 1.6.7 — optimallaştirilgan forma
Old/orqa natija maydonlari va printer çeckbox'i bitta `Natija va printer` GroupBox'iga birlaştirildi. Nusxalaş matni ixçam copy ikoniga almaştirildi, vertikal böşliqlar qisqartirildi va pastki tugmalar yuqoriga kötarildi. Forma 320×250 ölçamga tuşirildi.

## 1.6.6 — forma geometriyasi
Print Calculation oynasi 320×275 gaça ixçamlaştirildi. Barça GroupBox'lar 7 px taşqi çegara va bir xil 298 px kenglikka keltirildi; edit/copy juftliklari, vertikal oraliqlar va pastki tugmalar yagona grid böyiça tekislandi.

## 1.6.5 — GroupBox çild-control tuzatişi
Forma elementlari GroupBox ustiga emas, tegişli Frame içiga çild-control sifatida joylaştirildi. Hisob label'lari, natija maydonlari, nusxalaş tugmalari va printer çeckbox'i endi ramkalar ortida qolmaydi.

## 1.6.4 — ixçam Print Calculation formasi
Katta bannerli forma ixçam xizmat oynasiga almaştirildi. `Hisob`, `Old tomon`, `Orqa tomon` va `Printer` qismlari oddiy GroupBox bilan ajratildi; hisob label'lari ustma-ust joylaştirildi, natija maydonlari bir qatorga tuşirildi va forma ölçami 420×340 ga kamaytirildi. VBA kodlaşida buzilgan apostroflar oddiy apostrofga almaştirildi.

## 1.6.3 — Print Calculation runtime tuzatişi
`Set not permitted` xatosining haqiqiy manbasi dinamik ToggleButton uçun runtime'da ruxsat etilmagan `.SpecialEffect` xususiyati ekani Word smoke-test orqali aniqlandi. Xususiyat olib taşlandi; forma control'lari, dastlabki hisoblaş va natija maydonlariga yoziş haqiqiy Word jarayonida tekşirildi.

## 1.6.2 — Print Calculation forma tuzatişi
Dinamik TextBox maydonlariga fokusdan taşqarida `.Text` yozilganda yuzaga kelgan `Set not permitted` xatosi tuzatildi. Barça qiymatlar fokusga boğliq bölmagan `.Value` orqali boşqariladi.

## 1.6.1 — Word xavfsizlik ogohlantirişi
Installer qöşimça faylidagi internet blokini tozalaydi va joriy foydalanuvçining Word STARTUP papkasini aniq Trusted Location sifatida röyxatdan ötkazadi. Olib taşlaşda installer yaratgan işonç yozuvi ham öçiriladi.

## 1.6.0 — Kitobça hisoblagiç
`usta_hasan → print → Kitobcha hisoblagich` formasi old va orqa tomon uçun betlar ketma-ketligini avtomatik hisoblaydi. Bir tomonga 2, 4, 6, 8 yoki 16 bet tanlanadi. Forma jami/yakuniy sahifa, qöşiladigan böş sahifalar va kerakli qoğoz varaqlarini körsatadi; natijalarni alohida nusxalaş va yetişmagan betlarni hujjatga qöşiş mumkin.

## 1.5.7 — installer dizayni
Installer oynasi zamonaviyroq va ixçam vizual tizimga ötkazildi: ajratilgan sarlavha, versiya belgisi, Word lentasidagi yöl kartasi, örnatilganlik holati va aniq asosiy/ikkilamçi tugmalar qöşildi. Örnatiş funksiyalari özgarmadi.

## 1.5.6 tuzatişi
Katta hujjatdagi konvertatsiyalarni bitta Application.UndoRecord guruhiga yiğiş Word 16.0 ning wwlib.dll modulida 0xc0000005 xatosini keltirib çiqargani hujjat nusxasida takrorlandi.
Ikkala transliteratsiya makrosidan bu guruhlaş olib taşlandi. Almaştiriş jadvali özgarmadi.
Konvertatsiyani qaytariş uçun bir neça marta Ctrl+Z kerak bölişi mumkin.

## Işlaş tartibi
Tanlov bölsa faqat şu qism, tanlov bölmasa hujjatning asosiy matni işlanadi.
Header/footer va boşqa alohida matn qismlari avtomatik aylanib çiqilmaydi.
Lotin/kiril qoidalari foydalanuvçi bergan jadvallarga mos.
Kiril → Lotin: Ъ/ъ U+02BC (ʼ) ga aylanadi, Ь/ь olib taşlanadi; Э/э → E/e; Ы/ы va Щ/щ jadvalda bölmagani uçun saqlanadi.

## Sinovlar
30 ta kiçik funksional test qayta ötdi.
Haqiqiy hujjatning 18 888 belgili nusxasida ikkala konvertatsiya oxirigaça bajarildi.
Natijalar almaştiriş jadvallari bilan soliştirildi; 5 ta rasm saqlandi; natija saqlanib qayta oçildi.
Asl foydalanuvçi faylining ŞA-256 xeşi özgarmadi.
Hisobotlar: tests/word-test.log, tests/latin-test.log, tests/krill-test.log, tests/real-document-test.log.
Haqiqiy hujjat va uning sinov nusxalari tarqatiladigan ZIP içiga kiritilmaydi.

## Manbalar
src/KochishBelgisi.bas — belgi makrosi.
src/Latin_Krill.bas va src/Krill_Latin.bas — konvertatsiyalar.
src/customUI.xml — lenta va tugmalar.
src/Setup.cs — örnatuvçi kodi.
src/installer.ico — örnatuvçi ikonkasi.
src/word-2024.png — installer reklama maydonidagi Word belgisi va rang manbasi.
src/koçiş.png — Word tugmasi rasmi.
src/FormBelgi-designer.vbaProject.bin — forma dizayni bilan VBA loyihasi.
src/build_addin.py — Word qöşimçasini yiğiş.
Build.cmd — qöşimça va örnatuvçini qayta yiğiş.
src/package_release.py — zaxira va nazorat yiğindilari.
Normal.backup.dotm — asl forma dizayni uçun donor.

Qayta yiğiş: Python 3, oletools va .NET Framework C# kompilyatori kerak.
python -m pip install --target .tools oletools==0.60.2
Keyin Build.cmd ni işga tuşiring. Örnatiş uçun Python kerak emas.

## Qöşimça boşqaruvi
Qöşimça joriy foydalanuvçining Word STARTUP papkasiga örnatiladi.
Odatda: %APPDATA%\Microsoft\Word\STARTUP.
Olib taşlaş uçun Word'ni yoping va örnatuvçidagi Olib taşlaşni bosing.
Yangilaş zaxiralari: %LOCALAPPDATA%\usta_hasan\KochishBelgisi\backups.
Normal.dotm, Word.officeUI va xavfsizlik siyosatlari almaştirilmaydi.
usta_hasan RibbonX orqali yuklanadi; Customize Ribbon röyxatida körinmasligi asosiy lentada yöqligini bildirmaydi.

Hozirgi versiya: KochishBelgisi_Backup_1.6.10.zip. Eski versiyalarni qayta örnatmang.

## 1.5.6 — katta harflar va tutuq
ҚУЁШ → QUYOŞ; Қуёш → Quyoş; қуёш → quyoş.
Katta harfli sözlarda Ş/Ç/YO/YU/YA/TS ham töliq katta yoziladi.
MA’LUMOT → МАЪЛУМОТ; ma’lumot → маълумот.
Söz içidagi apostrof variantlari kirilda ъ/Ъ böladi. Taşqi qöştirnoqlar saqlanadi.
Ö/Ğ → Ў/Ғ alohida işlanadi. Kiril Ъ/ъ lotinda tutuq (U+02BC), Ў/Ғ esa U+02BB bilan yoziladi.
Matnning harf holati va tutuq joylari xotirada tekşiriladi; faqat kerakli belgilar Word orqali özgartiriladi.
Yangi misollar: tests/case-marks-test.log. Haqiqiy hujjat sarlavhalari ham tekşirildi.


## 1.5.6 — Köçiş belgisi formasi
Tanlov bölsa faqat tanlangan matn, aks holda hujjatning asosiy matni işlanadi.
Köçiş belgisi uçun avvalgi kiçik Belgilar hisoblagiçi qaytarildi. 10 ta yoki undan kam almaştirişda forma çiqmaydi. 10 tadan köp bölsa çiqadi; iş tugagaç yakuniy son 1 soniya körinib, forma yopiladi. Iş vaqtida X bosilsa navbatdagi işlov berişda makros töxtaydi; oldin bajarilgan almaştirişlar saqlanadi.
Matn xotirada tekşiriladi va faqat topilgan belgilar almaştiriladi.
Lotin/Kiril konvertatsiyalarida foiz va progress oynasi saqlangan.
Sinovlar: tests/form-threşold.log va tests/word-test.log.

## Konvertatsiya oynasi
Sarlavhada usta hasan va tanlangan yönaliş körsatiladi. Ixçam oynada bosqiç, foiz va hisoblagiç bor; pastdagi takroriy yozuv olib taşlandi. X orqali töxtatiş saqlangan. Köçiş belgisi formasi özgarmadi.

## 1.5.6 — havolali matn
Tanlov içidagi havolalar va Word maydonlari (ularning körinadigan matni ham) saqlanadi; atrofdagi oddiy matn konvertatsiya qilinadi. Havola manzili tahrirlanmaydi. Sinov: tests/fields-test.log.

1.5.6: Murakkab matn qismlarida Word pozitsiyalari böyiça xavfsiz zaxira öqiş qöşildi. Maqola.docx nusxasi 12,4 soniyada konvertatsiya qilindi; sarlavha, rasm soni va havola manzili tekşirildi.

ESKI ALFAVIT
Alt+3: Ö/ö → Oʻ/oʻ; Ğ/ğ → Gʻ/gʻ; Ş/ş → Sh/sh; Ç/ç → Ch/ch.
SH/CH bosh harfli söz kontekstida tiklanadi. Apostrof standart ʻ belgisi bilan yoziladi.
Bu matnni eski alifboga ögiradi; avvalgi apostrof turi aynan tiklanmaydi.
