TimeStudyPro — Login/Profile integration v1

ឯកសារក្នុង ZIP៖ index.html, style.css, script.js

អ្វីដែលបានធ្វើ៖
- ប្រើ index.html ដែលអ្នកបានផ្ញើជាគោល។
- ប្រើ script.js និង style.css ពី TimeStudyPro_3D_Feature_Removed_v1.zip ជាមូលដ្ឋាន។
- បញ្ចូល Login/Profile/User Management ទៅក្នុង script.js ដើម្បីមិនចាំបាច់បើក dashboard-analytics.js ឬ user-management.js ដាច់ដោយឡែក។
- បន្ថែម CSS សម្រាប់ Login/Profile/User Management នៅចុង style.css។
- មិនកែ logic របស់ Dashboard ក្នុង script.js។

គណនីដំបូងសម្រាប់សាកល្បង៖
Username: SOKRET11
Password: Admin123!

សំខាន់៖ កំណែនេះប្រើ localStorage ក្នុង Browser សម្រាប់សាកល្បងប៉ុណ្ណោះ មិនមែនជា authentication សុវត្ថិភាព server-side ទេ។


ការកែបន្ថែម v2៖
- កែផ្ទាំង Upload File > Action > Preview ឱ្យតូច និងទំនើបជាងមុន (ទទឹងអតិបរមា 380px)។
- បន្ថែម CSS សម្រាប់ responsive លើទូរស័ព្ទ។
- រក្សាមុខងារផ្សេងៗ និងឯកសារដើមទាំងអស់។


ការកែសម្រួល Preview (v3):
- ប្ដូរ Button Close ទៅជាសញ្ញា × នៅជ្រុងខាងស្ដាំខាងលើ។
- ដកប៊ូតុង Close ខាងក្រោមចេញ ដើម្បីកាត់បន្ថយកម្ពស់ប្រអប់។
- បង្រួមទំហំប្រអប់ និងចន្លោះខាងក្នុងឱ្យតូច ស្អាត និងសមស្របលើទូរស័ព្ទ។


ការកែលម្អ Upload File v4:
- File Card/Table UI និងព័ត៌មានទំហំ/ចំនួនឯកសារ។
- Drag & Drop និង Upload progress indicator។
- Search, Filter តាម .xls/.xlsx និង Sort តាមឈ្មោះ/កាលបរិច្ឆេទ/ទំហំ។
- Smart Preview តូចស្អាត មានប៊ូតុង × នៅជ្រុងខាងលើ។
- ការពារឈ្មោះ File ស្ទួន និងពិនិត្យមាតិកាស្ទួនតាម SHA-256 នៅ browser ដែលគាំទ្រ Web Crypto។
- ជ្រើសរើស File ច្រើន ដើម្បី Download ឬ Delete; ការលុបនៅតែគោរពសិទ្ធិ Admin/អ្នក Upload និងមានការបញ្ជាក់។
- Storage summary និងកំណត់ទំហំ File Excel អតិបរមា 10 MB។

ចំណាំ៖ ទិន្នន័យនៅតែរក្សាទុកតាម localStorage របស់ Browser ដូច Code ដើម។ Storage Summary បង្ហាញទំហំទិន្នន័យដែលរក្សាទុកប្រហែល មិនមែនទំហំ Disk/Server ទេ។ Browser អាចមានកម្រិត localStorage ហើយការប្រើ Web Crypto សម្រាប់ផ្ទៀងផ្ទាត់មាតិកា អាស្រ័យលើការគាំទ្ររបស់ Browser និង secure context។


ការកែលម្អ Upload File v5 — បន្ថែម 6 មុខងារ៖
- File Details Panel៖ ព័ត៌មាន File, អ្នក Upload, ទំហំ, ចំនួន Rows និង Version History។
- File Version History៖ អាចរក្សា Snapshot ដោយដៃ និង Restore កំណែ Snapshot បាន។ Snapshot រក្សាទុកក្នុង localStorage ដូចទិន្នន័យ File ដើម។
- Grid View / List View៖ ប្ដូរទម្រង់បង្ហាញ និងរក្សាទុកជម្រើសនៅ Browser។
- Recycle Bin៖ Delete ផ្លាស់ទៅ Recycle Bin ជាមុន មាន Restore និង Permanent Delete។
- Storage Dashboard៖ បង្ហាញចំនួន File, ទំហំដើមប្រហែល, File Type, File ធំបំផុត និងចំនួនក្នុង Recycle Bin។
- Permissions & Activity Log៖ បង្ហាញគោលការណ៍សិទ្ធិ និងកំណត់ត្រា Upload/Download/Delete/Restore។

ចំណាំសុវត្ថិភាព៖ មុខងារទាំងនេះដំណើរការនៅ Browser/localStorage។ សិទ្ធិមិនមែនជា Server-side security ទេ ហើយ Recycle Bin/Version History ប្រើទំហំផ្ទុកបន្ថែម។ គួររក្សា Backup នៃទិន្នន័យសំខាន់ៗ។

ការកែលម្អ Upload File v6 — Quick Filter តាមខែ/ឆ្នាំ៖
- បន្ថែមប៊ូតុង «ខែ / ឆ្នាំ» នៅជាប់ Dropdown តម្រៀប «ថ្មីបំផុត»។
- ជ្រើសខែ និងឆ្នាំ ដើម្បីបង្ហាញតែ File ដែលបាន Upload ក្នុងខែនោះ។
- អាចជ្រើស «បង្ហាញទាំងអស់» ដើម្បីសម្អាត Filter និងមើល File គ្រប់ខែវិញ។
- ចំនួន File ក្នុងចំណងជើងបញ្ជីបង្ហាញចំនួនដែលត្រូវនឹង Filter / ចំនួន File សរុប។
- រក្សា File Details/Version History, Grid/List, Recycle Bin, Storage Dashboard, Permissions/Activity Log និង Duplicate Detection ដែលមានស្រាប់។


មុខងារ Login/Admin បន្ថែមក្នុង v18:
- Admin Activity Notification: ចុចរូបកណ្ដឹងដើម្បីបើក/បិទផ្ទាំងជូនដំណឹង ដែលបង្ហាញចំនួន User មិនស្ទួនដែលបាន Upload, Delete/Move to Recycle Bin និង Download នៅថ្ងៃនេះ ព្រមទាំងរូប Profile ឈ្មោះ File សកម្មភាព និងម៉ោង។
- Live Active Users និង User Management: បង្ហាញរូប Profile ឬរូបសញ្ញា User ជាមួយឈ្មោះគណនី និងស្ថានភាព។ ផ្ទាំង Admin ត្រូវបានបង្រួមឱ្យស្រួលមើលលើទូរស័ព្ទ។
- Password: Admin អាច Reset Password ពី User Management បាន ប៉ុន្តែមិនបង្ហាញ Password ដែល User កំពុងប្រើ។
- កំណត់សម្គាល់សុវត្ថិភាព៖ វានៅតែជា Browser-local/localStorage demo ដូចកូដ v16។
- គណនី Admin ដំបូង (បើ LocalStorage មិនទាន់មានគណនី): SOKRET11 / Admin123!

v19 Update (Admin Activity UX)
- បន្ថែមលេខសារមិនទាន់អានលើសញ្ញាកណ្ដឹង Admin; ពេលបើកមើលសកម្មភាព លេខនឹងក្លាយជា 0 ហើយសកម្មភាពថ្មីនឹងរាប់ឡើងវិញ។
- ប៊ូតុង User Upload, User Delete និង User Download អាចចុចដើម្បីត្រងបញ្ជី User/រូប Profile តាមប្រភេទសកម្មភាព។
- បង្រួមផ្ទាំង Admin Activity ប្រហែល 25% លើ Desktop និងបន្តសម្របខ្លួនតាមអេក្រង់ទូរស័ព្ទ។

កំណែ v20៖
- កែទីតាំងចំនួនសារជូនដំណឹងឲ្យជាប់នៅលើសញ្ញាកណ្ដឹងផ្ទាល់។
- បង្រួមប្រអប់ Admin/User dialogs ប្រហែល 55% បន្ថែមពីទំហំកំណែមុន។
- បង្រួមប្រអប់សារជូនដំណឹងលើទូរស័ព្ទ និងកំណត់កម្ពស់ឲ្យតូចជាងមុន មាន scroll នៅពេលមានសារច្រើន។
- រក្សាមុខងារចាស់ និងស្ថិតិចុចត្រងតាម Upload/Delete/Download។

កំណែ v21៖ ផ្ទាំងស្អាត និងកាត់បន្ថយ Scroll
- ដក scroll ខាងក្នុងចេញពីប្រអប់ Dialog/Notification និងបញ្ជីប្រវត្តិ Upload ដែលពាក់ព័ន្ធ ដោយកែ Font-size, Padding, Gap និងទំហំតារាងឲ្យសមស្រប។
- ពង្រីកទំហំប្រអប់ដែលត្រូវការ ដើម្បីឲ្យមាតិកាបង្ហាញបានល្អជាងមុន ដោយមិនបង្ខំឲ្យអូសក្នុងប្រអប់។
- ផ្ទាំង Admin Activity បង្ហាញសកម្មភាពថ្មីៗបំផុត 5 ធាតុ ដើម្បីរក្សាផ្ទាំងឲ្យខ្លី និងគ្មាន scroll; ប៊ូតុង Upload/Delete/Download នៅដដែល។
- លើទូរស័ព្ទ កាត់បន្ថយទំហំអក្សរ និងចន្លោះ ហើយបត់អត្ថបទក្នុងតារាង ដើម្បីកាត់បន្ថយការអូសឆ្វេង/ស្តាំ។


V23 — Profile Photo Synchronization
- ពេល User រក្សាទុករូប Profile ថ្មី ឬដករូបចេញ ប្រព័ន្ធធ្វើបច្ចុប្បន្នភាពរូបនៅលើ Upload Files, Recycle Bin, File Version Snapshots និង Activity/Notification History ដែលស្គាល់ User នោះ។
- ធ្វើបច្ចុប្បន្នភាពឈ្មោះបង្ហាញរបស់អ្នក Upload ផងដែរ។
- រក្សាទុកមុខងារចាស់ និងទិន្នន័យឯកសារដើម។


============================================================
TimeStudyPro Smart Edition v24
============================================================
- Smart Profile Sync: ប្ដូររូប Profile ហើយធ្វើបច្ចុប្បន្នភាពឯកសារ/Activity references ដែលមានទិន្នន័យ User សម្គាល់បាន។
- Live User Status: បង្ហាញ Active/Offline និងពេលវេលា session ដែលមាន។
- Smart Activity Timeline: ស្វែងរកតាម User/File និងច្រោះតាមកាលបរិច្ឆេទ ព្រមទាំងប្រភេទ Upload/Delete/Download។
- User Security UI: Admin អាចបង្កើត/កែ User, Reset Password, បិទ/បើកគណនី និងកំណត់សិទ្ធិ Upload/Download/Delete own files។
- Compact responsive UI: កែចន្លោះ និងការបង្ហាញ filter សម្រាប់អេក្រង់តូច។

សុវត្ថិភាព៖ v24 នៅតែជាកម្មវិធី Browser-local ដែលប្រើ localStorage។ ការកំណត់សិទ្ធិ និង Password នៅខាង Client មិនអាចជំនួស Server-side authentication, password hashing និង database authorization បានទេ។ សម្រាប់ប្រើប្រាស់ពិតជាមួយ User ច្រើនឧបករណ៍ ត្រូវបន្ថែម Backend។


TimeStudyPro Smart Edition Plus v25 — additions
- Security & Login History: កត់ត្រា Login ជោគជ័យ/បរាជ័យ, Logout, ការបង្កើត/កែ/លុប User និងការកែសិទ្ធិគណនីក្នុង Browser-local audit log។
- Backup & Restore: Export ទិន្នន័យ localStorage ដែលពាក់ព័ន្ធទៅ JSON និង Restore ពី JSON បន្ទាប់ពីមានការបញ្ជាក់។ Backup អាចមានព័ត៌មានគណនី ដូច្នេះត្រូវរក្សាទុកឱ្យមានសុវត្ថិភាព។
- User Analytics Dashboard: បង្ហាញចំនួន User, Active User, Uploaded Files និងសកម្មភាពតាមប្រវត្តិដែលរក្សាទុកក្នុង Browser។
- Smart File Management: ស្វែងរក File តាមឈ្មោះ File/User និងសម្គាល់ Favorite។ Version History, Recycle Bin និង Storage Dashboard ដែលមានស្រាប់នៅតែរក្សាទុក។
- Smart Notification Center: មើលប្រវត្តិសកម្មភាព និង Mark all read។
- Personalized Workspace: កំណត់ Zoom, Font Size និង Light/Dark appearance ដែលរក្សាទុកក្នុង Browser នេះ។

សុវត្ថិភាព និងដែនកំណត់៖ v25 នៅតែជា Browser-local edition។ Audit log, permissions និងការចូលប្រើមិនអាចជំនួស Server authentication ឬការពារការកែ localStorage ដោយអ្នកមានសិទ្ធិចូល Browser បានទេ។ Backup JSON អាចមានទិន្នន័យសំខាន់ៗ; កុំចែករំលែកវា។ Restore គួរធ្វើក្រោយពេលទាញយក Backup បច្ចុប្បន្នរួច។ ទិន្នន័យ និងការកំណត់មិនសមកាលកម្មរវាងឧបករណ៍ទេ។


មុខងារថ្មី v26 — Data Ownership & Permissions
- User ទាំងអស់អាចមើល/Preview/Download និងប្រើទិន្នន័យរបស់អ្នកដទៃបាន។
- User អាចកែប្រែ ឬលុបបានតែ File/Time Study ដែលមាន Owner ID ត្រូវនឹងគណនីរបស់ខ្លួន។
- Admin អាចគ្រប់គ្រងទិន្នន័យទាំងអស់ រួមទាំង File History, Version History និង Recycle Bin។
- ទិន្នន័យចាស់ដែលគ្មាន Owner ID ត្រូវបានបិទសិទ្ធិកែប្រែ/លុបសម្រាប់ User ធម្មតា ដើម្បីការពារការលុបខុសម្ចាស់។
- ចំណាំ៖ កំណែនេះប្រើ localStorage ក្នុង Browser; ប្រសិនបើប្រើច្រើនឧបករណ៍/Server ត្រូវបន្ថែមការត្រួតពិនិត្យសិទ្ធិនៅ Server ផងដែរ។


============================================================
មុខងារថ្មី: Clone Day — Smart Merge
============================================================
- ជ្រើស File និង Source Day រួចបញ្ចូល Target Day មួយ ឬច្រើនដោយបំបែកជាសញ្ញាក្បៀស។
- Preview បង្ហាញចំនួន Process ត្រូវបន្ថែម និងចំនួនស្ទួនត្រូវរំលងមុន Confirm។
- Smart Merge រក្សាទិន្នន័យចាស់ និងបន្ថែមតែ Process ថ្មីប៉ុណ្ណោះ។
- មាន Confirm មុនអនុវត្ត និង Clone Report ក្រោយបញ្ចប់។
- ការកែប្រែអនុញ្ញាតតែម្ចាស់ File ឬ Admin ដោយរក្សាសិទ្ធិ User Management ដែលមានស្រាប់។
- Clone Day ប្រើទិន្នន័យ localStorage របស់កម្មវិធីនេះដូចមុខងារដើម។


ជំនាន់ v30 — Clone Day Source Picker A+B
- ជ្រើស File ជាមុន រួចបង្ហាញតែ Day ដែលមានក្នុង File នោះ។
- បញ្ជី Source Day បង្ហាញ Factory/Title, Section, Style និងចំនួន Process ដើម្បីឱ្យងាយសម្គាល់ថ្ងៃត្រឹមត្រូវ។
- មានប្រអប់ស្វែងរក Day/Section/Style និងសញ្ញាបញ្ជាក់ Source Day ដែលបានជ្រើស។
- រក្សា Smart Merge, Preview, ការពិនិត្យស្ទួន និងសិទ្ធិ User/Admin ដដែល។

ការកែថ្មី៖ Source Day Picker Modal
- ជ្រើស File ទិន្នន័យ រួចចុច “បើកតារាងជ្រើសថ្ងៃ” ដើម្បីបើកប្រអប់ជ្រើសថ្ងៃដើមពេញលេញ។
- ប្រអប់មានតារាង Day, Factory, Style, Section, Group, Process និងប៊ូតុងជ្រើស ព្រមទាំង Search។
- ចុច “ជ្រើសថ្ងៃនេះ” ដើម្បីកំណត់ថ្ងៃដើម ហើយបិទប្រអប់។
- បិទប្រអប់ដោយប៊ូតុង ×, ចុចក្រៅប្រអប់ ឬចុច Escape។

V20 targeted patch (based on TimeStudyPro_CloneDay_UltraGlass_Mobile_Menu_Fix_v19.zip)
- Added mobile-only CSS to reduce horizontal page overflow and scroll chaining.
- Kept horizontal scrolling inside wide tables.
- Set mobile text inputs to 16px to prevent iOS Safari auto-zoom on focus.
- Stabilized the fixed glass bottom navigation while scrolling.
- Desktop layout and existing JavaScript functions were not intentionally changed.

IMPORTANT — Login Sync status:
This package does NOT yet enable cross-device Login Sync. The existing user-management code stores users/passwords in browser localStorage, which cannot synchronize between devices. True sync requires a configured backend (for example Supabase Auth + database/RLS). The project URL and publishable/anon key and a reviewed database/auth migration are required before claiming Login Sync is active. Do not deploy this package as a completed Login Sync release.


ការកែលម្អ V34 — បិទ Pinch-to-Zoom លើទូរស័ព្ទ៖
- កែ viewport ដើម្បីស្នើឱ្យ Browser មិនអនុញ្ញាត page zoom ដោយម្រាមដៃពីរ។
- បន្ថែម touch-action pan-x/pan-y និងការការពារ multi-touch/gesture events។
- រក្សាការរមូរធម្មតាដោយម្រាមដៃមួយ និងការចុច Menu។
- មិនកែ logic មុខងារ TimeStudyPro និងមិនបន្ថែមមុខងារ Zoom ថ្មី។
ចំណាំ៖ Browser ឬការកំណត់ Accessibility មួយចំនួនអាចមិនគោរពការបិទ Zoom ទាំងស្រុងទេ។ ត្រូវសាកល្បងលើទូរស័ព្ទពិត។
