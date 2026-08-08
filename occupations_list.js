// ============================================================
// /occupations_list.js
// ------------------------------------------------------------
// A short reference list of occupation titles, offered as
// suggestions on the Occupation Title field.
//
// This is a TYPING AID, nothing more. The field stays completely
// free-text: the user can ignore every suggestion and type any
// occupation they like. Nothing here constrains, validates or
// rejects what they enter.
//
// SCOPE — why 100 and not 3,000:
// ESCO publishes ~3,039 occupations in 28 languages including
// Arabic, and it remains the right source for a future expansion.
// But most of it is non-technical, and burying a TVET facilitator
// under irrelevant suggestions is worse than a shorter list that
// hits. The 100 below are weighted toward the trades, technical
// and service occupations that vocational programmes actually
// cover.
//
// A list that is too short backfires: the user types a letter,
// finds nothing, and learns to ignore the dropdown permanently —
// including later, when it would have helped.
//
// THE `isco` FIELD IS DORMANT ON PURPOSE.
// It is not displayed, not written to appState, and not exported.
// It is recorded now only because writing it costs nothing while
// the row is being authored, whereas adding it later would mean
// revisiting all 100 rows by hand.
//
// !! DO NOT SURFACE THESE CODES WITHOUT VERIFYING THEM FIRST !!
// They were assigned from knowledge of the ISCO-08 structure, not
// copied from the ILO index. Codes on the boundary between two
// unit groups are the likeliest to be wrong. Before any feature
// displays or exports them, check each against ISCO-08 Volume II
// (the official index of occupational titles). Keeping them
// invisible today means an error here cannot reach a report or an
// accreditation body.
//
// ar / fr are filled in now for the same reason: the row is being
// written once either way, and the app is heading toward Arabic
// and French. Switching the displayed language is one constant.
// ============================================================

// Which language the suggestions are shown in: 'en' | 'ar' | 'fr'.
// The UI is currently English. When the i18n work resumes this can
// read from the active locale instead of being fixed here.
export const OCCUPATION_LIST_LANG = 'en';

export const OCCUPATIONS = [
  // ── Automotive & transport ──────────────────────────────────
  { isco: '7231', en: 'Automotive Service Technician',        ar: 'فني صيانة سيارات',                fr: 'Technicien de maintenance automobile' },
  { isco: '7231', en: 'Heavy Vehicle Mechanic',               ar: 'ميكانيكي مركبات ثقيلة',            fr: 'Mécanicien de véhicules lourds' },
  { isco: '7232', en: 'Aircraft Maintenance Technician',      ar: 'فني صيانة طائرات',                 fr: 'Technicien de maintenance aéronautique' },
  { isco: '7233', en: 'Agricultural Machinery Mechanic',      ar: 'ميكانيكي آلات زراعية',             fr: 'Mécanicien de machines agricoles' },
  { isco: '7213', en: 'Auto Body Repairer',                   ar: 'فني إصلاح هياكل السيارات',         fr: 'Carrossier automobile' },
  { isco: '7132', en: 'Vehicle Spray Painter',                ar: 'فني دهان السيارات',                fr: 'Peintre en carrosserie' },
  { isco: '8332', en: 'Heavy Truck Driver',                   ar: 'سائق شاحنة ثقيلة',                 fr: 'Conducteur de poids lourd' },
  { isco: '8343', en: 'Crane Operator',                       ar: 'مشغّل رافعة',                      fr: 'Grutier' },
  { isco: '8342', en: 'Earthmoving Equipment Operator',       ar: 'مشغّل معدات حفر ونقل',             fr: 'Conducteur d\u2019engins de terrassement' },
  { isco: '4323', en: 'Logistics and Warehouse Clerk',        ar: 'كاتب لوجستيات ومخازن',             fr: 'Agent logistique et magasinier' },

  // ── Electrical & electronics ────────────────────────────────
  { isco: '7411', en: 'Building Electrician',                 ar: 'كهربائي مبانٍ',                    fr: 'Électricien du bâtiment' },
  { isco: '7412', en: 'Industrial Electrician',               ar: 'كهربائي صناعي',                    fr: 'Électricien industriel' },
  { isco: '7413', en: 'Power Line Technician',                ar: 'فني خطوط نقل الكهرباء',            fr: 'Monteur de lignes électriques' },
  { isco: '7421', en: 'Electronics Technician',               ar: 'فني إلكترونيات',                   fr: 'Technicien en électronique' },
  { isco: '3113', en: 'Electrical Engineering Technician',    ar: 'فني هندسة كهربائية',               fr: 'Technicien en génie électrique' },
  { isco: '3114', en: 'Electronics Engineering Technician',   ar: 'فني هندسة إلكترونية',              fr: 'Technicien en génie électronique' },
  { isco: '7422', en: 'Telecommunications Technician',        ar: 'فني اتصالات',                      fr: 'Technicien en télécommunications' },
  { isco: '3522', en: 'Network Infrastructure Technician',    ar: 'فني بنية الشبكات',                 fr: 'Technicien réseaux' },
  { isco: '7421', en: 'Solar PV Installer',                   ar: 'فني تركيب الطاقة الشمسية',         fr: 'Installateur photovoltaïque' },
  { isco: '7412', en: 'Industrial Automation Technician',     ar: 'فني أتمتة صناعية',                 fr: 'Technicien en automatisation industrielle' },

  // ── Mechanical, metal & manufacturing ───────────────────────
  { isco: '7233', en: 'Industrial Machinery Mechanic',        ar: 'ميكانيكي آلات صناعية',             fr: 'Mécanicien de machines industrielles' },
  { isco: '7212', en: 'Welder',                               ar: 'لحّام',                            fr: 'Soudeur' },
  { isco: '7213', en: 'Sheet Metal Worker',                   ar: 'فني تشكيل الصفائح المعدنية',       fr: 'Tôlier' },
  { isco: '7214', en: 'Structural Steel Fabricator',          ar: 'فني تركيب الهياكل المعدنية',       fr: 'Charpentier métallique' },
  { isco: '7223', en: 'CNC Machine Operator',                 ar: 'مشغّل آلات CNC',                   fr: 'Opérateur sur machine CNC' },
  { isco: '7222', en: 'Toolmaker',                            ar: 'صانع عدد وقوالب',                  fr: 'Outilleur' },
  { isco: '7211', en: 'Foundry Moulder',                      ar: 'فني سباكة المعادن',                fr: 'Mouleur de fonderie' },
  { isco: '3115', en: 'Mechanical Engineering Technician',    ar: 'فني هندسة ميكانيكية',              fr: 'Technicien en génie mécanique' },
  { isco: '7126', en: 'Plumber',                              ar: 'سبّاك',                            fr: 'Plombier' },
  { isco: '7127', en: 'Air Conditioning and Refrigeration Technician', ar: 'فني تكييف وتبريد',      fr: 'Technicien en climatisation et froid' },

  // ── Construction & built environment ────────────────────────
  { isco: '7115', en: 'Carpenter',                            ar: 'نجّار',                            fr: 'Charpentier' },
  { isco: '7115', en: 'Furniture Maker',                      ar: 'صانع أثاث',                        fr: 'Ébéniste' },
  { isco: '7112', en: 'Bricklayer',                           ar: 'بنّاء',                            fr: 'Maçon' },
  { isco: '7114', en: 'Concrete Worker',                      ar: 'فني أعمال الخرسانة',               fr: 'Ouvrier du béton' },
  { isco: '7122', en: 'Floor and Tile Layer',                 ar: 'فني تبليط وأرضيات',                fr: 'Carreleur' },
  { isco: '7123', en: 'Plasterer',                            ar: 'مليّس',                            fr: 'Plâtrier' },
  { isco: '7131', en: 'Painter and Decorator',                ar: 'دهّان وديكور',                     fr: 'Peintre décorateur' },
  { isco: '7125', en: 'Glazier',                              ar: 'فني زجاج',                         fr: 'Vitrier' },
  { isco: '7121', en: 'Roofer',                               ar: 'فني أسطح ومظلات',                  fr: 'Couvreur' },
  { isco: '3112', en: 'Civil Engineering Technician',         ar: 'فني هندسة مدنية',                  fr: 'Technicien en génie civil' },
  { isco: '3118', en: 'Draughtsperson / CAD Technician',      ar: 'رسّام فني / فني CAD',              fr: 'Dessinateur / Technicien CAO' },
  { isco: '3112', en: 'Quantity Surveyor Assistant',          ar: 'مساعد مساح كميات',                 fr: 'Assistant métreur' },
  { isco: '3112', en: 'Site Supervisor',                      ar: 'مشرف موقع',                        fr: 'Chef de chantier' },
  { isco: '7119', en: 'Scaffolder',                           ar: 'فني سقالات',                       fr: 'Échafaudeur' },
  { isco: '7133', en: 'Building Insulation Installer',        ar: 'فني عزل المباني',                  fr: 'Installateur d\u2019isolation' },

  // ── ICT & digital ───────────────────────────────────────────
  { isco: '3512', en: 'IT Support Technician',                ar: 'فني دعم تقنية المعلومات',          fr: 'Technicien support informatique' },
  { isco: '2512', en: 'Software Developer',                   ar: 'مطوّر برمجيات',                    fr: 'Développeur logiciel' },
  { isco: '2513', en: 'Web Developer',                        ar: 'مطوّر مواقع',                      fr: 'Développeur web' },
  { isco: '2523', en: 'Computer Network Administrator',       ar: 'مسؤول شبكات حاسوب',                fr: 'Administrateur réseau' },
  { isco: '2522', en: 'Systems Administrator',                ar: 'مسؤول أنظمة',                      fr: 'Administrateur systèmes' },
  { isco: '2529', en: 'Cybersecurity Technician',             ar: 'فني أمن سيبراني',                  fr: 'Technicien en cybersécurité' },
  { isco: '2519', en: 'Database Technician',                  ar: 'فني قواعد بيانات',                 fr: 'Technicien base de données' },
  { isco: '2166', en: 'Graphic Designer',                     ar: 'مصمم جرافيك',                      fr: 'Graphiste' },
  { isco: '2166', en: 'Multimedia Designer',                  ar: 'مصمم وسائط متعددة',                fr: 'Concepteur multimédia' },
  { isco: '3521', en: 'Video and Audio Technician',           ar: 'فني صوت وصورة',                    fr: 'Technicien audiovisuel' },

  // ── Health & care ───────────────────────────────────────────
  { isco: '3221', en: 'Nursing Associate',                    ar: 'مساعد تمريض',                      fr: 'Infirmier auxiliaire' },
  { isco: '2221', en: 'Registered Nurse',                     ar: 'ممرض مسجّل',                       fr: 'Infirmier diplômé' },
  { isco: '3212', en: 'Medical Laboratory Technician',        ar: 'فني مختبر طبي',                    fr: 'Technicien de laboratoire médical' },
  { isco: '3211', en: 'Radiology Technician',                 ar: 'فني أشعة',                         fr: 'Technicien en radiologie' },
  { isco: '3213', en: 'Pharmacy Technician',                  ar: 'فني صيدلة',                        fr: 'Préparateur en pharmacie' },
  { isco: '3251', en: 'Dental Assistant',                     ar: 'مساعد طبيب أسنان',                 fr: 'Assistant dentaire' },
  { isco: '3214', en: 'Dental Laboratory Technician',         ar: 'فني مختبر أسنان',                  fr: 'Prothésiste dentaire' },
  { isco: '3214', en: 'Prosthetics and Orthotics Technician', ar: 'فني أطراف صناعية وتقويم',          fr: 'Technicien en orthoprothèse' },
  { isco: '3255', en: 'Physiotherapy Assistant',              ar: 'مساعد علاج طبيعي',                 fr: 'Assistant kinésithérapeute' },
  { isco: '5321', en: 'Home Care Aide',                       ar: 'مقدّم رعاية منزلية',               fr: 'Aide à domicile' },
  { isco: '3258', en: 'Emergency Medical Technician',         ar: 'فني طوارئ طبية',                   fr: 'Technicien ambulancier' },

  // ── Hospitality, food & tourism ─────────────────────────────
  { isco: '3434', en: 'Chef',                                 ar: 'رئيس طهاة',                        fr: 'Chef de cuisine' },
  { isco: '5120', en: 'Cook',                                 ar: 'طبّاخ',                            fr: 'Cuisinier' },
  { isco: '7512', en: 'Baker and Pastry Maker',               ar: 'خبّاز وصانع حلويات',               fr: 'Boulanger-pâtissier' },
  { isco: '7511', en: 'Butcher',                              ar: 'جزّار',                            fr: 'Boucher' },
  { isco: '5131', en: 'Waiter',                               ar: 'نادل',                             fr: 'Serveur' },
  { isco: '5132', en: 'Barista and Bartender',                ar: 'باريستا وساقي مشروبات',            fr: 'Barista et barman' },
  { isco: '4224', en: 'Hotel Receptionist',                   ar: 'موظف استقبال فندقي',               fr: 'Réceptionniste d\u2019hôtel' },
  { isco: '5151', en: 'Housekeeping Supervisor',              ar: 'مشرف خدمات الإقامة',               fr: 'Gouvernant d\u2019hôtel' },
  { isco: '4221', en: 'Travel Consultant',                    ar: 'مستشار سفر',                       fr: 'Conseiller en voyages' },
  { isco: '5113', en: 'Tour Guide',                           ar: 'مرشد سياحي',                       fr: 'Guide touristique' },

  // ── Personal services & crafts ──────────────────────────────
  { isco: '5141', en: 'Hairdresser',                          ar: 'مصفف شعر',                         fr: 'Coiffeur' },
  { isco: '5142', en: 'Beauty Therapist',                     ar: 'أخصائي تجميل',                     fr: 'Esthéticien' },
  { isco: '7531', en: 'Tailor and Dressmaker',                ar: 'خيّاط',                            fr: 'Tailleur-couturier' },
  { isco: '7533', en: 'Sewing Machine Operator',              ar: 'مشغّل ماكينة خياطة',               fr: 'Opérateur de machine à coudre' },
  { isco: '7536', en: 'Shoemaker and Leather Worker',         ar: 'صانع أحذية ومشغولات جلدية',        fr: 'Cordonnier et maroquinier' },
  { isco: '7314', en: 'Ceramics and Pottery Maker',           ar: 'صانع خزف وفخار',                   fr: 'Céramiste' },
  { isco: '7313', en: 'Jeweller and Goldsmith',               ar: 'صائغ ذهب ومجوهرات',                fr: 'Bijoutier-joaillier' },
  { isco: '7311', en: 'Watch and Instrument Repairer',        ar: 'فني إصلاح ساعات وأجهزة دقيقة',     fr: 'Horloger-réparateur' },
  { isco: '7317', en: 'Traditional Handicraft Maker',         ar: 'صانع حرف تقليدية',                 fr: 'Artisan traditionnel' },
  { isco: '5164', en: 'Animal Care Worker',                   ar: 'عامل رعاية حيوانات',               fr: 'Soigneur animalier' },

  // ── Agriculture, food processing & environment ──────────────
  { isco: '6111', en: 'Crop Production Worker',               ar: 'عامل إنتاج نباتي',                 fr: 'Ouvrier de production végétale' },
  { isco: '6121', en: 'Livestock Producer',                   ar: 'مربّي ماشية',                      fr: 'Éleveur de bétail' },
  { isco: '6113', en: 'Horticulturist and Nursery Grower',    ar: 'فني بستنة ومشاتل',                 fr: 'Horticulteur-pépiniériste' },
  { isco: '3142', en: 'Agricultural Technician',              ar: 'فني زراعي',                        fr: 'Technicien agricole' },
  { isco: '7511', en: 'Food Processing Technician',           ar: 'فني تصنيع أغذية',                  fr: 'Technicien agroalimentaire' },
  { isco: '3141', en: 'Food Quality Control Technician',      ar: 'فني ضبط جودة الأغذية',             fr: 'Technicien contrôle qualité alimentaire' },
  { isco: '3257', en: 'Environmental Health Inspector',       ar: 'مفتش صحة بيئية',                   fr: 'Inspecteur en santé environnementale' },
  { isco: '3132', en: 'Water Treatment Plant Operator',       ar: 'مشغّل محطة معالجة مياه',           fr: 'Opérateur de station de traitement des eaux' },

  // ── Business, administration & support ──────────────────────
  { isco: '3313', en: 'Accounting Technician',                ar: 'فني محاسبة',                       fr: 'Technicien comptable' },
  { isco: '4110', en: 'Office Administrator',                 ar: 'إداري مكتب',                       fr: 'Employé administratif' },
  { isco: '3341', en: 'Administrative Supervisor',            ar: 'مشرف إداري',                       fr: 'Superviseur administratif' },
  { isco: '3322', en: 'Sales Representative',                 ar: 'مندوب مبيعات',                     fr: 'Représentant commercial' },
  { isco: '5223', en: 'Retail Salesperson',                   ar: 'بائع تجزئة',                       fr: 'Vendeur en magasin' },
  { isco: '4222', en: 'Call Centre Agent',                    ar: 'موظف مركز اتصال',                  fr: 'Téléconseiller' },
  { isco: '3343', en: 'Executive Secretary',                  ar: 'سكرتير تنفيذي',                    fr: 'Secrétaire de direction' },
  { isco: '2423', en: 'Human Resources Officer',              ar: 'أخصائي موارد بشرية',               fr: 'Chargé des ressources humaines' },

  // ── Education, safety & security ────────────────────────────
  { isco: '2320', en: 'Vocational Training Instructor',       ar: 'مدرّب تدريب مهني',                 fr: 'Formateur en enseignement professionnel' },
  { isco: '2359', en: 'Curriculum Developer',                 ar: 'مطوّر مناهج',                      fr: 'Concepteur de programmes de formation' },
  { isco: '3257', en: 'Occupational Safety and Health Officer', ar: 'أخصائي سلامة وصحة مهنية',       fr: 'Responsable santé et sécurité au travail' },
  { isco: '5414', en: 'Security Guard',                       ar: 'حارس أمن',                         fr: 'Agent de sécurité' },
  { isco: '5411', en: 'Firefighter',                          ar: 'رجل إطفاء',                        fr: 'Sapeur-pompier' },
];

// Titles in the configured language, de-duplicated and sorted.
// Sorted because <datalist> renders the list in source order, and an
// unsorted dropdown reads as arbitrary to anyone scanning it.
export function getOccupationTitles(lang = OCCUPATION_LIST_LANG) {
  const key = ['en', 'ar', 'fr'].includes(lang) ? lang : 'en';
  return [...new Set(OCCUPATIONS.map(o => o[key]).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, key === 'ar' ? 'ar' : key));
}

// Populates the <datalist> feeding the Occupation Title input.
//
// A native <datalist> is used rather than a custom dropdown: no
// library, works offline, keyboard and screen-reader accessible for
// free, and — most importantly — it cannot block free text. The
// trade-off is that match behaviour differs between browsers (Chrome
// matches any substring, some match only the start) and it cannot be
// styled. Both are acceptable for a typing aid; neither would be for
// a control the user had to rely on.
export function initOccupationSuggestions(lang) {
  const input = document.getElementById('occupationTitle');
  const list  = document.getElementById('occupationTitleList');
  if (!input || !list) return;

  list.innerHTML = getOccupationTitles(lang)
    .map(t => `<option value="${t.replace(/"/g, '&quot;')}"></option>`)
    .join('');

  input.setAttribute('list', 'occupationTitleList');
}
