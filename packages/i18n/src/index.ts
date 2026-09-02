export type Lang =
  | 'zh'
  | 'en'
  | 'ja'
  | 'ko'
  | 'fr'
  | 'de'
  | 'es'
  | 'th'
  | 'id'
  | 'ru'
  | 'ar'
  | 'pt'
  | 'it'
  | 'pl'
  | 'nl'
  | 'ms'
  | 'he'
  | 'hi'
  | 'zh-TW'

export const LANGS: readonly Lang[] = [
  'zh',
  'en',
  'ja',
  'ko',
  'fr',
  'de',
  'es',
  'th',
  'id',
  'ru',
  'ar',
  'pt',
  'it',
  'pl',
  'nl',
  'ms',
  'he',
  'hi',
  'zh-TW',
]

export function isLang(value: unknown): value is Lang {
  return typeof value === 'string' && (LANGS as readonly string[]).includes(value)
}

/** map a raw locale string ('zh-CN', 'zh-Hans', 'ja-JP', 'ko-KR', …) to a supported Lang */
export function normalizeLang(raw: string | null | undefined): Lang {
  const value = raw?.trim().toLowerCase()
  if (!value) return 'en'
  // traditional-script Chinese variants must win over the generic 'zh' prefix
  if (/^zh[-_](tw|hk|mo|hant)/.test(value)) return 'zh-TW'
  for (const lang of LANGS) {
    if (lang !== 'en' && lang !== 'zh-TW' && value.startsWith(lang)) return lang
  }
  // 'in' is the legacy ISO code for Indonesian still reported by some systems
  if (/^in\b/.test(value) || /^in[-_]/.test(value)) return 'id'
  // 'iw' is the legacy ISO code for Hebrew
  if (/^iw\b/.test(value) || /^iw[-_]/.test(value)) return 'he'
  return 'en'
}

/** Elevated raw Office confirmation heading. The exact program remains visible in the proposal. */
export const RAW_OFFICE_CONFIRMATION: Readonly<Record<Lang, string>> = Object.freeze({
  zh: '高权限 Office 编辑：请逐项核对后确认',
  en: 'Elevated Office edit: review every change before confirming',
  ja: '高権限の Office 編集：すべての変更を確認してください',
  ko: '고급 Office 편집: 확인 전에 모든 변경 사항을 검토하세요',
  fr: 'Modification Office avancée : vérifiez chaque changement avant de confirmer',
  de: 'Erweiterte Office-Bearbeitung: Prüfen Sie jede Änderung vor der Bestätigung',
  es: 'Edición avanzada de Office: revisa cada cambio antes de confirmar',
  th: 'การแก้ไข Office ขั้นสูง: ตรวจสอบทุกการเปลี่ยนแปลงก่อนยืนยัน',
  id: 'Pengeditan Office tingkat lanjut: tinjau setiap perubahan sebelum mengonfirmasi',
  ru: 'Расширенное редактирование Office: проверьте каждое изменение перед подтверждением',
  ar: 'تحرير Office بصلاحيات مرتفعة: راجع كل تغيير قبل التأكيد',
  pt: 'Edição avançada do Office: revise cada alteração antes de confirmar',
  it: 'Modifica Office avanzata: controlla ogni modifica prima di confermare',
  pl: 'Zaawansowana edycja Office: sprawdź każdą zmianę przed potwierdzeniem',
  nl: 'Office-bewerking met verhoogde rechten: controleer elke wijziging vóór bevestiging',
  ms: 'Suntingan Office lanjutan: semak setiap perubahan sebelum mengesahkan',
  he: 'עריכת Office בהרשאה מוגברת: יש לבדוק כל שינוי לפני האישור',
  hi: 'उन्नत Office संपादन: पुष्टि से पहले हर बदलाव की समीक्षा करें',
  'zh-TW': '高權限 Office 編輯：請逐項檢查後確認',
})

export function translateRawOfficeConfirmation(lang: Lang): string {
  return RAW_OFFICE_CONFIRMATION[lang]
}

const mutationEnglish = {
  title: 'Confirm document change',
  operation: 'Operation',
  target: 'Target',
  scope: 'Scope',
  count: 'Items affected',
  warning:
    'Review this request carefully. WisWork will apply it as one bounded transaction only after you confirm.',
  reject: 'Reject',
  confirm: 'Confirm change',
  'operation.insert': 'Insert',
  'operation.replace': 'Replace',
  'operation.delete': 'Delete',
  'operation.format': 'Format',
  'operation.restructure': 'Restructure',
  'operation.compile': 'Compile',
  'target.document': 'Document',
  'target.selection': 'Selection',
  'target.blocks': 'Blocks',
  'target.cells': 'Cells',
  'target.sheet': 'Sheet',
  'target.slides': 'Slides',
  'target.elements': 'Elements',
  'target.project-files': 'Project files',
  'scope.single': 'Single item',
  'scope.selection': 'Current selection',
  'scope.bounded-set': 'Bounded set',
  'scope.whole-document': 'Whole document',
} as const

export type EnhancedMutationConfirmationKey = keyof typeof mutationEnglish
type MutationConfirmationStrings = Record<EnhancedMutationConfirmationKey, string>

/** Complete informed-consent copy for bounded Enhanced mutations. */
export const enhancedMutationConfirmationStrings: Readonly<
  Record<Lang, MutationConfirmationStrings>
> = Object.freeze({
  zh: {
    title: '确认文档更改',
    operation: '操作类型',
    target: '目标',
    scope: '影响范围',
    count: '影响数量',
    warning: '请仔细检查本次操作。只有在你确认后，WisWork 才会将其作为一个受限事务执行。',
    reject: '拒绝',
    confirm: '确认更改',
    'operation.insert': '插入',
    'operation.replace': '替换',
    'operation.delete': '删除',
    'operation.format': '格式调整',
    'operation.restructure': '结构调整',
    'operation.compile': '编译',
    'target.document': '文档',
    'target.selection': '选区',
    'target.blocks': '内容块',
    'target.cells': '单元格',
    'target.sheet': '工作表',
    'target.slides': '幻灯片',
    'target.elements': '页面元素',
    'target.project-files': '项目文件',
    'scope.single': '单个对象',
    'scope.selection': '当前选区',
    'scope.bounded-set': '受限集合',
    'scope.whole-document': '整个文档',
  },
  en: mutationEnglish,
  ja: {
    title: '文書の変更を確認',
    operation: '操作',
    target: '対象',
    scope: '範囲',
    count: '影響する項目数',
    warning:
      'この操作をよく確認してください。確認後にのみ、WisWork が限定された単一トランザクションとして適用します。',
    reject: '拒否',
    confirm: '変更を確認',
    'operation.insert': '挿入',
    'operation.replace': '置換',
    'operation.delete': '削除',
    'operation.format': '書式設定',
    'operation.restructure': '再構成',
    'operation.compile': 'コンパイル',
    'target.document': '文書',
    'target.selection': '選択範囲',
    'target.blocks': 'ブロック',
    'target.cells': 'セル',
    'target.sheet': 'シート',
    'target.slides': 'スライド',
    'target.elements': '要素',
    'target.project-files': 'プロジェクトファイル',
    'scope.single': '単一項目',
    'scope.selection': '現在の選択範囲',
    'scope.bounded-set': '限定された集合',
    'scope.whole-document': '文書全体',
  },
  ko: {
    title: '문서 변경 확인',
    operation: '작업',
    target: '대상',
    scope: '범위',
    count: '영향받는 항목',
    warning:
      '이 요청을 주의 깊게 검토하세요. 확인한 후에만 WisWork가 제한된 단일 트랜잭션으로 적용합니다.',
    reject: '거부',
    confirm: '변경 확인',
    'operation.insert': '삽입',
    'operation.replace': '바꾸기',
    'operation.delete': '삭제',
    'operation.format': '서식 지정',
    'operation.restructure': '재구성',
    'operation.compile': '컴파일',
    'target.document': '문서',
    'target.selection': '선택 영역',
    'target.blocks': '블록',
    'target.cells': '셀',
    'target.sheet': '시트',
    'target.slides': '슬라이드',
    'target.elements': '요소',
    'target.project-files': '프로젝트 파일',
    'scope.single': '단일 항목',
    'scope.selection': '현재 선택 영역',
    'scope.bounded-set': '제한된 집합',
    'scope.whole-document': '전체 문서',
  },
  fr: {
    title: 'Confirmer la modification du document',
    operation: 'Opération',
    target: 'Cible',
    scope: 'Portée',
    count: 'Éléments concernés',
    warning:
      'Examinez attentivement cette demande. WisWork ne l’appliquera comme transaction limitée qu’après votre confirmation.',
    reject: 'Refuser',
    confirm: 'Confirmer la modification',
    'operation.insert': 'Insérer',
    'operation.replace': 'Remplacer',
    'operation.delete': 'Supprimer',
    'operation.format': 'Mettre en forme',
    'operation.restructure': 'Restructurer',
    'operation.compile': 'Compiler',
    'target.document': 'Document',
    'target.selection': 'Sélection',
    'target.blocks': 'Blocs',
    'target.cells': 'Cellules',
    'target.sheet': 'Feuille',
    'target.slides': 'Diapositives',
    'target.elements': 'Éléments',
    'target.project-files': 'Fichiers du projet',
    'scope.single': 'Élément unique',
    'scope.selection': 'Sélection actuelle',
    'scope.bounded-set': 'Ensemble limité',
    'scope.whole-document': 'Document entier',
  },
  de: {
    title: 'Dokumentänderung bestätigen',
    operation: 'Vorgang',
    target: 'Ziel',
    scope: 'Umfang',
    count: 'Betroffene Elemente',
    warning:
      'Prüfen Sie diese Anfrage sorgfältig. WisWork wendet sie erst nach Ihrer Bestätigung als einzelne begrenzte Transaktion an.',
    reject: 'Ablehnen',
    confirm: 'Änderung bestätigen',
    'operation.insert': 'Einfügen',
    'operation.replace': 'Ersetzen',
    'operation.delete': 'Löschen',
    'operation.format': 'Formatieren',
    'operation.restructure': 'Neu strukturieren',
    'operation.compile': 'Kompilieren',
    'target.document': 'Dokument',
    'target.selection': 'Auswahl',
    'target.blocks': 'Blöcke',
    'target.cells': 'Zellen',
    'target.sheet': 'Tabelle',
    'target.slides': 'Folien',
    'target.elements': 'Elemente',
    'target.project-files': 'Projektdateien',
    'scope.single': 'Einzelnes Element',
    'scope.selection': 'Aktuelle Auswahl',
    'scope.bounded-set': 'Begrenzte Menge',
    'scope.whole-document': 'Gesamtes Dokument',
  },
  es: {
    title: 'Confirmar cambio del documento',
    operation: 'Operación',
    target: 'Destino',
    scope: 'Alcance',
    count: 'Elementos afectados',
    warning:
      'Revisa esta solicitud con atención. WisWork solo la aplicará como una transacción limitada después de tu confirmación.',
    reject: 'Rechazar',
    confirm: 'Confirmar cambio',
    'operation.insert': 'Insertar',
    'operation.replace': 'Reemplazar',
    'operation.delete': 'Eliminar',
    'operation.format': 'Dar formato',
    'operation.restructure': 'Reestructurar',
    'operation.compile': 'Compilar',
    'target.document': 'Documento',
    'target.selection': 'Selección',
    'target.blocks': 'Bloques',
    'target.cells': 'Celdas',
    'target.sheet': 'Hoja',
    'target.slides': 'Diapositivas',
    'target.elements': 'Elementos',
    'target.project-files': 'Archivos del proyecto',
    'scope.single': 'Elemento único',
    'scope.selection': 'Selección actual',
    'scope.bounded-set': 'Conjunto limitado',
    'scope.whole-document': 'Documento completo',
  },
  th: {
    title: 'ยืนยันการเปลี่ยนแปลงเอกสาร',
    operation: 'การดำเนินการ',
    target: 'เป้าหมาย',
    scope: 'ขอบเขต',
    count: 'รายการที่ได้รับผลกระทบ',
    warning:
      'โปรดตรวจสอบคำขอนี้อย่างรอบคอบ WisWork จะนำไปใช้เป็นธุรกรรมที่จำกัดหนึ่งรายการหลังจากคุณยืนยันเท่านั้น',
    reject: 'ปฏิเสธ',
    confirm: 'ยืนยันการเปลี่ยนแปลง',
    'operation.insert': 'แทรก',
    'operation.replace': 'แทนที่',
    'operation.delete': 'ลบ',
    'operation.format': 'จัดรูปแบบ',
    'operation.restructure': 'ปรับโครงสร้าง',
    'operation.compile': 'คอมไพล์',
    'target.document': 'เอกสาร',
    'target.selection': 'ส่วนที่เลือก',
    'target.blocks': 'บล็อก',
    'target.cells': 'เซลล์',
    'target.sheet': 'แผ่นงาน',
    'target.slides': 'สไลด์',
    'target.elements': 'องค์ประกอบ',
    'target.project-files': 'ไฟล์โครงการ',
    'scope.single': 'รายการเดียว',
    'scope.selection': 'ส่วนที่เลือกปัจจุบัน',
    'scope.bounded-set': 'ชุดที่จำกัด',
    'scope.whole-document': 'ทั้งเอกสาร',
  },
  id: {
    title: 'Konfirmasi perubahan dokumen',
    operation: 'Operasi',
    target: 'Target',
    scope: 'Cakupan',
    count: 'Item terdampak',
    warning:
      'Tinjau permintaan ini dengan saksama. WisWork hanya akan menerapkannya sebagai satu transaksi terbatas setelah Anda mengonfirmasi.',
    reject: 'Tolak',
    confirm: 'Konfirmasi perubahan',
    'operation.insert': 'Sisipkan',
    'operation.replace': 'Ganti',
    'operation.delete': 'Hapus',
    'operation.format': 'Format',
    'operation.restructure': 'Susun ulang',
    'operation.compile': 'Kompilasi',
    'target.document': 'Dokumen',
    'target.selection': 'Pilihan',
    'target.blocks': 'Blok',
    'target.cells': 'Sel',
    'target.sheet': 'Lembar',
    'target.slides': 'Slide',
    'target.elements': 'Elemen',
    'target.project-files': 'File proyek',
    'scope.single': 'Satu item',
    'scope.selection': 'Pilihan saat ini',
    'scope.bounded-set': 'Kumpulan terbatas',
    'scope.whole-document': 'Seluruh dokumen',
  },
  ru: {
    title: 'Подтвердите изменение документа',
    operation: 'Операция',
    target: 'Объект',
    scope: 'Область',
    count: 'Затронуто элементов',
    warning:
      'Внимательно проверьте запрос. WisWork применит его как одну ограниченную транзакцию только после подтверждения.',
    reject: 'Отклонить',
    confirm: 'Подтвердить изменение',
    'operation.insert': 'Вставить',
    'operation.replace': 'Заменить',
    'operation.delete': 'Удалить',
    'operation.format': 'Форматировать',
    'operation.restructure': 'Изменить структуру',
    'operation.compile': 'Скомпилировать',
    'target.document': 'Документ',
    'target.selection': 'Выделение',
    'target.blocks': 'Блоки',
    'target.cells': 'Ячейки',
    'target.sheet': 'Лист',
    'target.slides': 'Слайды',
    'target.elements': 'Элементы',
    'target.project-files': 'Файлы проекта',
    'scope.single': 'Один элемент',
    'scope.selection': 'Текущее выделение',
    'scope.bounded-set': 'Ограниченный набор',
    'scope.whole-document': 'Весь документ',
  },
  ar: {
    title: 'تأكيد تغيير المستند',
    operation: 'العملية',
    target: 'الهدف',
    scope: 'النطاق',
    count: 'العناصر المتأثرة',
    warning: 'راجع هذا الطلب بعناية. لن يطبّقه WisWork كمعاملة واحدة محدودة إلا بعد تأكيدك.',
    reject: 'رفض',
    confirm: 'تأكيد التغيير',
    'operation.insert': 'إدراج',
    'operation.replace': 'استبدال',
    'operation.delete': 'حذف',
    'operation.format': 'تنسيق',
    'operation.restructure': 'إعادة هيكلة',
    'operation.compile': 'ترجمة',
    'target.document': 'المستند',
    'target.selection': 'التحديد',
    'target.blocks': 'الكتل',
    'target.cells': 'الخلايا',
    'target.sheet': 'ورقة العمل',
    'target.slides': 'الشرائح',
    'target.elements': 'العناصر',
    'target.project-files': 'ملفات المشروع',
    'scope.single': 'عنصر واحد',
    'scope.selection': 'التحديد الحالي',
    'scope.bounded-set': 'مجموعة محدودة',
    'scope.whole-document': 'المستند بالكامل',
  },
  pt: {
    title: 'Confirmar alteração do documento',
    operation: 'Operação',
    target: 'Destino',
    scope: 'Escopo',
    count: 'Itens afetados',
    warning:
      'Revise esta solicitação com atenção. O WisWork só a aplicará como uma transação limitada após sua confirmação.',
    reject: 'Rejeitar',
    confirm: 'Confirmar alteração',
    'operation.insert': 'Inserir',
    'operation.replace': 'Substituir',
    'operation.delete': 'Excluir',
    'operation.format': 'Formatar',
    'operation.restructure': 'Reestruturar',
    'operation.compile': 'Compilar',
    'target.document': 'Documento',
    'target.selection': 'Seleção',
    'target.blocks': 'Blocos',
    'target.cells': 'Células',
    'target.sheet': 'Planilha',
    'target.slides': 'Slides',
    'target.elements': 'Elementos',
    'target.project-files': 'Arquivos do projeto',
    'scope.single': 'Item único',
    'scope.selection': 'Seleção atual',
    'scope.bounded-set': 'Conjunto limitado',
    'scope.whole-document': 'Documento inteiro',
  },
  it: {
    title: 'Conferma modifica al documento',
    operation: 'Operazione',
    target: 'Destinazione',
    scope: 'Ambito',
    count: 'Elementi interessati',
    warning:
      'Esamina attentamente la richiesta. WisWork la applicherà come singola transazione limitata solo dopo la conferma.',
    reject: 'Rifiuta',
    confirm: 'Conferma modifica',
    'operation.insert': 'Inserisci',
    'operation.replace': 'Sostituisci',
    'operation.delete': 'Elimina',
    'operation.format': 'Formatta',
    'operation.restructure': 'Ristruttura',
    'operation.compile': 'Compila',
    'target.document': 'Documento',
    'target.selection': 'Selezione',
    'target.blocks': 'Blocchi',
    'target.cells': 'Celle',
    'target.sheet': 'Foglio',
    'target.slides': 'Diapositive',
    'target.elements': 'Elementi',
    'target.project-files': 'File del progetto',
    'scope.single': 'Elemento singolo',
    'scope.selection': 'Selezione corrente',
    'scope.bounded-set': 'Insieme limitato',
    'scope.whole-document': 'Intero documento',
  },
  pl: {
    title: 'Potwierdź zmianę dokumentu',
    operation: 'Operacja',
    target: 'Cel',
    scope: 'Zakres',
    count: 'Elementy objęte zmianą',
    warning:
      'Dokładnie sprawdź to żądanie. WisWork zastosuje je jako jedną ograniczoną transakcję dopiero po potwierdzeniu.',
    reject: 'Odrzuć',
    confirm: 'Potwierdź zmianę',
    'operation.insert': 'Wstaw',
    'operation.replace': 'Zastąp',
    'operation.delete': 'Usuń',
    'operation.format': 'Formatuj',
    'operation.restructure': 'Zmień strukturę',
    'operation.compile': 'Kompiluj',
    'target.document': 'Dokument',
    'target.selection': 'Zaznaczenie',
    'target.blocks': 'Bloki',
    'target.cells': 'Komórki',
    'target.sheet': 'Arkusz',
    'target.slides': 'Slajdy',
    'target.elements': 'Elementy',
    'target.project-files': 'Pliki projektu',
    'scope.single': 'Jeden element',
    'scope.selection': 'Bieżące zaznaczenie',
    'scope.bounded-set': 'Ograniczony zestaw',
    'scope.whole-document': 'Cały dokument',
  },
  nl: {
    title: 'Documentwijziging bevestigen',
    operation: 'Bewerking',
    target: 'Doel',
    scope: 'Bereik',
    count: 'Betrokken items',
    warning:
      'Controleer dit verzoek zorgvuldig. WisWork past het pas na uw bevestiging toe als één begrensde transactie.',
    reject: 'Weigeren',
    confirm: 'Wijziging bevestigen',
    'operation.insert': 'Invoegen',
    'operation.replace': 'Vervangen',
    'operation.delete': 'Verwijderen',
    'operation.format': 'Opmaken',
    'operation.restructure': 'Herstructureren',
    'operation.compile': 'Compileren',
    'target.document': 'Document',
    'target.selection': 'Selectie',
    'target.blocks': 'Blokken',
    'target.cells': 'Cellen',
    'target.sheet': 'Werkblad',
    'target.slides': 'Dia’s',
    'target.elements': 'Elementen',
    'target.project-files': 'Projectbestanden',
    'scope.single': 'Eén item',
    'scope.selection': 'Huidige selectie',
    'scope.bounded-set': 'Begrensde set',
    'scope.whole-document': 'Hele document',
  },
  ms: {
    title: 'Sahkan perubahan dokumen',
    operation: 'Operasi',
    target: 'Sasaran',
    scope: 'Skop',
    count: 'Item terjejas',
    warning:
      'Semak permintaan ini dengan teliti. WisWork hanya akan melaksanakannya sebagai satu transaksi terhad selepas anda mengesahkan.',
    reject: 'Tolak',
    confirm: 'Sahkan perubahan',
    'operation.insert': 'Sisip',
    'operation.replace': 'Ganti',
    'operation.delete': 'Padam',
    'operation.format': 'Format',
    'operation.restructure': 'Susun semula',
    'operation.compile': 'Kompil',
    'target.document': 'Dokumen',
    'target.selection': 'Pilihan',
    'target.blocks': 'Blok',
    'target.cells': 'Sel',
    'target.sheet': 'Helaian',
    'target.slides': 'Slaid',
    'target.elements': 'Elemen',
    'target.project-files': 'Fail projek',
    'scope.single': 'Satu item',
    'scope.selection': 'Pilihan semasa',
    'scope.bounded-set': 'Set terhad',
    'scope.whole-document': 'Seluruh dokumen',
  },
  he: {
    title: 'אישור שינוי במסמך',
    operation: 'פעולה',
    target: 'יעד',
    scope: 'היקף',
    count: 'פריטים שיושפעו',
    warning: 'יש לבדוק בקשה זו בקפידה. WisWork יחיל אותה כעסקה מוגבלת אחת רק לאחר האישור שלך.',
    reject: 'דחייה',
    confirm: 'אישור השינוי',
    'operation.insert': 'הוספה',
    'operation.replace': 'החלפה',
    'operation.delete': 'מחיקה',
    'operation.format': 'עיצוב',
    'operation.restructure': 'שינוי מבנה',
    'operation.compile': 'הידור',
    'target.document': 'מסמך',
    'target.selection': 'בחירה',
    'target.blocks': 'מקטעים',
    'target.cells': 'תאים',
    'target.sheet': 'גיליון',
    'target.slides': 'שקופיות',
    'target.elements': 'רכיבים',
    'target.project-files': 'קובצי פרויקט',
    'scope.single': 'פריט יחיד',
    'scope.selection': 'הבחירה הנוכחית',
    'scope.bounded-set': 'קבוצה מוגבלת',
    'scope.whole-document': 'המסמך כולו',
  },
  hi: {
    title: 'दस्तावेज़ परिवर्तन की पुष्टि करें',
    operation: 'कार्रवाई',
    target: 'लक्ष्य',
    scope: 'दायरा',
    count: 'प्रभावित आइटम',
    warning:
      'इस अनुरोध की सावधानी से समीक्षा करें। आपकी पुष्टि के बाद ही WisWork इसे एक सीमित लेन-देन के रूप में लागू करेगा।',
    reject: 'अस्वीकार करें',
    confirm: 'परिवर्तन की पुष्टि करें',
    'operation.insert': 'सम्मिलित करें',
    'operation.replace': 'बदलें',
    'operation.delete': 'हटाएँ',
    'operation.format': 'फ़ॉर्मैट करें',
    'operation.restructure': 'पुनर्गठित करें',
    'operation.compile': 'कंपाइल करें',
    'target.document': 'दस्तावेज़',
    'target.selection': 'चयन',
    'target.blocks': 'ब्लॉक',
    'target.cells': 'सेल',
    'target.sheet': 'शीट',
    'target.slides': 'स्लाइड',
    'target.elements': 'तत्व',
    'target.project-files': 'प्रोजेक्ट फ़ाइलें',
    'scope.single': 'एक आइटम',
    'scope.selection': 'वर्तमान चयन',
    'scope.bounded-set': 'सीमित समूह',
    'scope.whole-document': 'पूरा दस्तावेज़',
  },
  'zh-TW': {
    title: '確認文件變更',
    operation: '操作類型',
    target: '目標',
    scope: '影響範圍',
    count: '影響數量',
    warning: '請仔細檢查本次操作。只有在你確認後，WisWork 才會將其作為一個受限交易執行。',
    reject: '拒絕',
    confirm: '確認變更',
    'operation.insert': '插入',
    'operation.replace': '取代',
    'operation.delete': '刪除',
    'operation.format': '格式調整',
    'operation.restructure': '結構調整',
    'operation.compile': '編譯',
    'target.document': '文件',
    'target.selection': '選取範圍',
    'target.blocks': '內容區塊',
    'target.cells': '儲存格',
    'target.sheet': '工作表',
    'target.slides': '投影片',
    'target.elements': '頁面元素',
    'target.project-files': '專案檔案',
    'scope.single': '單一物件',
    'scope.selection': '目前選取範圍',
    'scope.bounded-set': '受限集合',
    'scope.whole-document': '整份文件',
  },
})

export function translateEnhancedMutationConfirmation(
  lang: Lang,
  key: EnhancedMutationConfirmationKey,
): string {
  return enhancedMutationConfirmationStrings[lang][key]
}

const enhancedEnglish = {
  label: 'Enhanced mode',
  download: 'Download',
  download_again: 'Download again',
  restart_required: 'Restart WisWork to apply',
  unavailable: 'Not supported on this device',
  blocked_by_policy: 'Codex Enhanced is disabled by policy',
  failed_safe: 'Codex Enhanced failed safely; switch to Standard or restart to retry',
  install_required: 'Install required before use',
  optional_download: 'Optional download required',
  switch_standard: 'Switch to Standard mode',
  enable_after_restart: 'Enable after restart',
  remove: 'Remove optional component',
  standard: 'Standard mode',
  enhanced: 'Codex Enhanced',
  rollback: 'Rollback available',
}

/** Complete, shared lifecycle copy used by Shell and all paired editor surfaces. */
export const enhancedModeStrings = defineStrings({
  zh: {
    label: '增强模式',
    download: '下载',
    download_again: '重新下载',
    restart_required: '重启 WisWork 后生效',
    unavailable: '此设备暂不支持',
    blocked_by_policy: '当前策略未开放 Codex 增强模式',
    failed_safe: 'Codex 增强模式已安全停止；请切换到标准模式或重启后重试',
    install_required: '使用前需要安装',
    optional_download: '需要下载可选组件',
    switch_standard: '切换到标准模式',
    enable_after_restart: '启用（重启后生效）',
    remove: '移除可选组件',
    standard: '标准模式',
    enhanced: '增强模式',
    rollback: '可回滚',
  },
  en: enhancedEnglish,
  ja: {
    label: 'Codex 拡張モード',
    download: 'ダウンロード',
    download_again: '再ダウンロード',
    restart_required: 'WisWork の再起動後に適用されます',
    unavailable: 'このデバイスでは利用できません',
    blocked_by_policy: 'ポリシーにより Codex 拡張モードは無効です',
    failed_safe: 'Codex 拡張モードは安全に停止しました。標準モードへ切り替えるか再起動してください',
    install_required: '使用前にインストールが必要です',
    optional_download: 'オプション部品のダウンロードが必要です',
    switch_standard: '標準モードに切り替える',
    enable_after_restart: '再起動後に有効化',
    remove: 'オプション部品を削除',
    standard: '標準モード',
    enhanced: 'Codex 拡張モード',
    rollback: '元に戻せます',
  },
  ko: {
    label: 'Codex 향상 모드',
    download: '다운로드',
    download_again: '다시 다운로드',
    restart_required: 'WisWork를 다시 시작하면 적용됩니다',
    unavailable: '이 장치에서는 지원되지 않습니다',
    blocked_by_policy: '정책에 의해 Codex 향상 모드가 비활성화되었습니다',
    failed_safe:
      'Codex 향상 모드가 안전하게 중지되었습니다. 표준 모드로 전환하거나 다시 시작하세요',
    install_required: '사용 전에 설치해야 합니다',
    optional_download: '선택 구성 요소 다운로드가 필요합니다',
    switch_standard: '표준 모드로 전환',
    enable_after_restart: '재시작 후 활성화',
    remove: '선택 구성 요소 제거',
    standard: '표준 모드',
    enhanced: 'Codex 향상 모드',
    rollback: '롤백 가능',
  },
  fr: {
    label: 'Codex amélioré',
    download: 'Télécharger',
    download_again: 'Télécharger à nouveau',
    restart_required: 'Redémarrez WisWork pour appliquer',
    unavailable: 'Non pris en charge sur cet appareil',
    blocked_by_policy: 'Codex amélioré est désactivé par la politique',
    failed_safe: 'Codex amélioré s’est arrêté en sécurité ; passez en mode Standard ou redémarrez',
    install_required: 'Installation requise avant utilisation',
    optional_download: 'Téléchargement facultatif requis',
    switch_standard: 'Passer en mode Standard',
    enable_after_restart: 'Activer après redémarrage',
    remove: 'Supprimer le composant facultatif',
    standard: 'Mode Standard',
    enhanced: 'Codex amélioré',
    rollback: 'Restauration disponible',
  },
  de: {
    label: 'Codex Erweitert',
    download: 'Herunterladen',
    download_again: 'Erneut herunterladen',
    restart_required: 'WisWork neu starten, um anzuwenden',
    unavailable: 'Auf diesem Gerät nicht unterstützt',
    blocked_by_policy: 'Codex Erweitert ist durch Richtlinie deaktiviert',
    failed_safe: 'Codex Erweitert wurde sicher beendet; zu Standard wechseln oder neu starten',
    install_required: 'Installation vor Verwendung erforderlich',
    optional_download: 'Optionaler Download erforderlich',
    switch_standard: 'Zum Standardmodus wechseln',
    enable_after_restart: 'Nach Neustart aktivieren',
    remove: 'Optionale Komponente entfernen',
    standard: 'Standardmodus',
    enhanced: 'Codex Erweitert',
    rollback: 'Rollback verfügbar',
  },
  es: {
    label: 'Codex mejorado',
    download: 'Descargar',
    download_again: 'Descargar de nuevo',
    restart_required: 'Reinicia WisWork para aplicar',
    unavailable: 'No compatible con este dispositivo',
    blocked_by_policy: 'Codex mejorado está desactivado por la política',
    failed_safe: 'Codex mejorado se detuvo de forma segura; cambia a Estándar o reinicia',
    install_required: 'Instalación requerida antes de usar',
    optional_download: 'Se requiere una descarga opcional',
    switch_standard: 'Cambiar al modo Estándar',
    enable_after_restart: 'Activar después de reiniciar',
    remove: 'Quitar componente opcional',
    standard: 'Modo Estándar',
    enhanced: 'Codex mejorado',
    rollback: 'Reversión disponible',
  },
  th: {
    label: 'Codex ขั้นสูง',
    download: 'ดาวน์โหลด',
    download_again: 'ดาวน์โหลดอีกครั้ง',
    restart_required: 'รีสตาร์ท WisWork เพื่อใช้งาน',
    unavailable: 'อุปกรณ์นี้ไม่รองรับ',
    blocked_by_policy: 'นโยบายปิดใช้งาน Codex ขั้นสูง',
    failed_safe: 'Codex ขั้นสูงหยุดอย่างปลอดภัย โปรดสลับเป็นโหมดมาตรฐานหรือรีสตาร์ท',
    install_required: 'ต้องติดตั้งก่อนใช้งาน',
    optional_download: 'ต้องดาวน์โหลดส่วนประกอบเสริม',
    switch_standard: 'สลับเป็นโหมดมาตรฐาน',
    enable_after_restart: 'เปิดใช้หลังรีสตาร์ท',
    remove: 'ลบส่วนประกอบเสริม',
    standard: 'โหมดมาตรฐาน',
    enhanced: 'Codex ขั้นสูง',
    rollback: 'ย้อนกลับได้',
  },
  id: {
    label: 'Codex Ditingkatkan',
    download: 'Unduh',
    download_again: 'Unduh lagi',
    restart_required: 'Mulai ulang WisWork untuk menerapkan',
    unavailable: 'Tidak didukung di perangkat ini',
    blocked_by_policy: 'Codex Ditingkatkan dinonaktifkan oleh kebijakan',
    failed_safe: 'Codex Ditingkatkan berhenti dengan aman; beralih ke Standar atau mulai ulang',
    install_required: 'Instalasi diperlukan sebelum digunakan',
    optional_download: 'Unduhan opsional diperlukan',
    switch_standard: 'Beralih ke mode Standar',
    enable_after_restart: 'Aktifkan setelah mulai ulang',
    remove: 'Hapus komponen opsional',
    standard: 'Mode Standar',
    enhanced: 'Codex Ditingkatkan',
    rollback: 'Pemulihan tersedia',
  },
  ru: {
    label: 'Codex Enhanced',
    download: 'Скачать',
    download_again: 'Скачать снова',
    restart_required: 'Перезапустите WisWork для применения',
    unavailable: 'Не поддерживается на этом устройстве',
    blocked_by_policy: 'Codex Enhanced отключён политикой',
    failed_safe:
      'Codex Enhanced безопасно остановлен; выберите стандартный режим или перезапустите',
    install_required: 'Перед использованием требуется установка',
    optional_download: 'Требуется загрузить дополнительный компонент',
    switch_standard: 'Перейти в стандартный режим',
    enable_after_restart: 'Включить после перезапуска',
    remove: 'Удалить дополнительный компонент',
    standard: 'Стандартный режим',
    enhanced: 'Codex Enhanced',
    rollback: 'Доступен откат',
  },
  ar: {
    label: 'Codex المحسّن',
    download: 'تنزيل',
    download_again: 'إعادة التنزيل',
    restart_required: 'أعد تشغيل WisWork للتطبيق',
    unavailable: 'غير مدعوم على هذا الجهاز',
    blocked_by_policy: 'تم تعطيل Codex المحسّن بواسطة السياسة',
    failed_safe: 'توقف Codex المحسّن بأمان؛ انتقل إلى الوضع القياسي أو أعد التشغيل',
    install_required: 'يلزم التثبيت قبل الاستخدام',
    optional_download: 'يلزم تنزيل المكوّن الاختياري',
    switch_standard: 'التبديل إلى الوضع القياسي',
    enable_after_restart: 'تمكين بعد إعادة التشغيل',
    remove: 'إزالة المكوّن الاختياري',
    standard: 'الوضع القياسي',
    enhanced: 'Codex المحسّن',
    rollback: 'التراجع متاح',
  },
  pt: {
    label: 'Codex Avançado',
    download: 'Baixar',
    download_again: 'Baixar novamente',
    restart_required: 'Reinicie o WisWork para aplicar',
    unavailable: 'Não compatível com este dispositivo',
    blocked_by_policy: 'Codex Avançado está desativado pela política',
    failed_safe: 'Codex Avançado parou com segurança; mude para Padrão ou reinicie',
    install_required: 'Instalação necessária antes do uso',
    optional_download: 'Download opcional necessário',
    switch_standard: 'Mudar para o modo Padrão',
    enable_after_restart: 'Ativar após reiniciar',
    remove: 'Remover componente opcional',
    standard: 'Modo Padrão',
    enhanced: 'Codex Avançado',
    rollback: 'Reversão disponível',
  },
  it: {
    label: 'Codex Avanzato',
    download: 'Scarica',
    download_again: 'Scarica di nuovo',
    restart_required: 'Riavvia WisWork per applicare',
    unavailable: 'Non supportato su questo dispositivo',
    blocked_by_policy: 'Codex Avanzato è disabilitato dai criteri',
    failed_safe: 'Codex Avanzato si è arrestato in sicurezza; passa a Standard o riavvia',
    install_required: 'Installazione richiesta prima dell’uso',
    optional_download: 'Download opzionale richiesto',
    switch_standard: 'Passa alla modalità Standard',
    enable_after_restart: 'Abilita dopo il riavvio',
    remove: 'Rimuovi componente opzionale',
    standard: 'Modalità Standard',
    enhanced: 'Codex Avanzato',
    rollback: 'Ripristino disponibile',
  },
  pl: {
    label: 'Codex Rozszerzony',
    download: 'Pobierz',
    download_again: 'Pobierz ponownie',
    restart_required: 'Uruchom WisWork ponownie, aby zastosować',
    unavailable: 'Brak obsługi na tym urządzeniu',
    blocked_by_policy: 'Codex Rozszerzony jest wyłączony przez zasady',
    failed_safe:
      'Codex Rozszerzony zatrzymał się bezpiecznie; przełącz na Standard lub uruchom ponownie',
    install_required: 'Przed użyciem wymagana jest instalacja',
    optional_download: 'Wymagane pobranie opcjonalnego składnika',
    switch_standard: 'Przełącz na tryb Standard',
    enable_after_restart: 'Włącz po ponownym uruchomieniu',
    remove: 'Usuń opcjonalny składnik',
    standard: 'Tryb Standard',
    enhanced: 'Codex Rozszerzony',
    rollback: 'Wycofanie dostępne',
  },
  nl: {
    label: 'Codex Uitgebreid',
    download: 'Downloaden',
    download_again: 'Opnieuw downloaden',
    restart_required: 'Start WisWork opnieuw om toe te passen',
    unavailable: 'Niet ondersteund op dit apparaat',
    blocked_by_policy: 'Codex Uitgebreid is uitgeschakeld door beleid',
    failed_safe: 'Codex Uitgebreid is veilig gestopt; schakel naar Standaard of start opnieuw',
    install_required: 'Installatie vereist voor gebruik',
    optional_download: 'Optionele download vereist',
    switch_standard: 'Naar Standaardmodus',
    enable_after_restart: 'Inschakelen na herstart',
    remove: 'Optioneel onderdeel verwijderen',
    standard: 'Standaardmodus',
    enhanced: 'Codex Uitgebreid',
    rollback: 'Terugdraaien beschikbaar',
  },
  ms: {
    label: 'Codex Dipertingkat',
    download: 'Muat turun',
    download_again: 'Muat turun semula',
    restart_required: 'Mulakan semula WisWork untuk digunakan',
    unavailable: 'Tidak disokong pada peranti ini',
    blocked_by_policy: 'Codex Dipertingkat dilumpuhkan oleh dasar',
    failed_safe:
      'Codex Dipertingkat berhenti dengan selamat; tukar ke Standard atau mulakan semula',
    install_required: 'Pemasangan diperlukan sebelum digunakan',
    optional_download: 'Muat turun pilihan diperlukan',
    switch_standard: 'Tukar ke mod Standard',
    enable_after_restart: 'Dayakan selepas mula semula',
    remove: 'Buang komponen pilihan',
    standard: 'Mod Standard',
    enhanced: 'Codex Dipertingkat',
    rollback: 'Pemulihan tersedia',
  },
  he: {
    label: 'Codex משופר',
    download: 'הורדה',
    download_again: 'הורדה מחדש',
    restart_required: 'יש להפעיל מחדש את WisWork כדי להחיל',
    unavailable: 'לא נתמך במכשיר זה',
    blocked_by_policy: 'Codex משופר מושבת לפי מדיניות',
    failed_safe: 'Codex משופר נעצר בבטחה; עברו למצב רגיל או הפעילו מחדש',
    install_required: 'נדרשת התקנה לפני השימוש',
    optional_download: 'נדרשת הורדה אופציונלית',
    switch_standard: 'מעבר למצב רגיל',
    enable_after_restart: 'הפעלה לאחר אתחול',
    remove: 'הסרת רכיב אופציונלי',
    standard: 'מצב רגיל',
    enhanced: 'Codex משופר',
    rollback: 'שחזור זמין',
  },
  hi: {
    label: 'Codex उन्नत',
    download: 'डाउनलोड करें',
    download_again: 'फिर डाउनलोड करें',
    restart_required: 'लागू करने के लिए WisWork पुनः आरंभ करें',
    unavailable: 'इस डिवाइस पर समर्थित नहीं',
    blocked_by_policy: 'नीति ने Codex उन्नत को बंद किया है',
    failed_safe: 'Codex उन्नत सुरक्षित रूप से रुका; मानक मोड चुनें या पुनः आरंभ करें',
    install_required: 'उपयोग से पहले इंस्टॉल करना आवश्यक है',
    optional_download: 'वैकल्पिक डाउनलोड आवश्यक है',
    switch_standard: 'मानक मोड पर जाएँ',
    enable_after_restart: 'पुनः आरंभ के बाद सक्षम करें',
    remove: 'वैकल्पिक घटक हटाएँ',
    standard: 'मानक मोड',
    enhanced: 'Codex उन्नत',
    rollback: 'वापसी उपलब्ध है',
  },
  'zh-TW': {
    label: 'Codex 增強模式',
    download: '下載',
    download_again: '重新下載',
    restart_required: '重新啟動 WisWork 後生效',
    unavailable: '此裝置暫不支援',
    blocked_by_policy: '目前政策未開放 Codex 增強模式',
    failed_safe: 'Codex 增強模式已安全停止；請切換至標準模式或重新啟動',
    install_required: '使用前需要安裝',
    optional_download: '需要下載選用元件',
    switch_standard: '切換至標準模式',
    enable_after_restart: '啟用（重新啟動後生效）',
    remove: '移除選用元件',
    standard: '標準模式',
    enhanced: 'Codex 增強模式',
    rollback: '可復原',
  },
})

export type EnhancedModeStringKey = keyof typeof enhancedModeStrings.zh
export function translateEnhancedMode(lang: Lang, key: EnhancedModeStringKey): string {
  return enhancedModeStrings[lang][key]
}

const HTML_LANGS: Record<Lang, string> = {
  zh: 'zh-CN',
  en: 'en-US',
  ja: 'ja-JP',
  ko: 'ko-KR',
  fr: 'fr-FR',
  de: 'de-DE',
  es: 'es-ES',
  th: 'th-TH',
  id: 'id-ID',
  ru: 'ru-RU',
  ar: 'ar-SA',
  pt: 'pt-BR',
  it: 'it-IT',
  pl: 'pl-PL',
  nl: 'nl-NL',
  ms: 'ms-MY',
  he: 'he-IL',
  hi: 'hi-IN',
  'zh-TW': 'zh-TW',
}

/** BCP-47 tag for document.documentElement.lang (drives CSS :lang() and Chromium's per-language font fallback) */
export function htmlLang(lang: Lang): string {
  return HTML_LANGS[lang]
}

// ---- platform-native shortcut hints ----
// Dictionaries write shortcut hints in Mac notation (⌘S, ⇧⌘Z, ⌘+Click); on
// Windows/Linux every translated string is rewritten to Ctrl/Alt/Shift form.

const MAC_KEY_NAMES: Record<string, string> = {
  '⌫': 'Backspace',
  '⌦': 'Delete',
  '⏎': 'Enter',
  '↩': 'Enter',
  '␣': 'Space',
}

const HAS_MAC_SYMBOL = /[⌘⌃⌥⇧⌫⌦⏎↩␣]/
const CHORD = /([⌘⌃⌥⇧]+)(F\d{1,2}|[A-Za-z0-9±=`'\\,./;[\]\-←↑→↓⌫⌦⏎↩␣]|\+)?/g

function chordToWin(mods: string, key: string | undefined): string {
  const parts: string[] = []
  if (mods.includes('⌘') || mods.includes('⌃')) parts.push('Ctrl')
  if (mods.includes('⌥')) parts.push('Alt')
  if (mods.includes('⇧')) parts.push('Shift')
  if (key) parts.push(MAC_KEY_NAMES[key] ?? key)
  return parts.join('+')
}

/** rewrite Mac shortcut notation in a UI string to Windows/Linux form (pure) */
export function macShortcutsToWin(text: string): string {
  if (!HAS_MAC_SYMBOL.test(text)) return text
  return text
    .replace(/⌘\/(?=\p{L}{2})/gu, '') // "⌘/Ctrl+Enter" dual-platform listings: keep the Ctrl side
    .replace(CHORD, (_m, mods: string, key: string | undefined) =>
      key === '+' ? `${chordToWin(mods, undefined)}+` : chordToWin(mods, key),
    )
    .replace(/[⌫⌦⏎↩␣]/g, (glyph) => MAC_KEY_NAMES[glyph] ?? glyph)
}

const IS_MAC = (() => {
  const g = globalThis as {
    navigator?: { platform?: string }
    process?: { platform?: string }
  }
  if (g.navigator?.platform) return /mac/i.test(g.navigator.platform)
  return g.process?.platform === 'darwin'
})()

/** platform-aware shortcut display: identity on macOS */
export const platformShortcuts: (text: string) => string = IS_MAC
  ? (text) => text
  : macShortcutsToWin

export type Params = Record<string, string | number>

/** fill {name} placeholders; unknown placeholders are left as-is */
export function format(template: string, params?: Params): string {
  if (!params) return template
  return template.replace(/\{(\w+)\}/g, (match, name: string) =>
    name in params ? String(params[name]) : match,
  )
}

/** per-language dictionaries; zh defines the key set, all others must match it */
export type LangDicts<D extends Record<string, string>> = { zh: D } & {
  [L in Exclude<Lang, 'zh'>]: Record<keyof D, string>
}

/**
 * Identity helper for dictionary shards: keeps literal key inference while
 * type-checking that every other language covers exactly the zh key set.
 */
export function defineStrings<D extends Record<string, string>>(dicts: LangDicts<D>): LangDicts<D> {
  return dicts
}

/** User-visible presentation verification states. Codes remain locale-neutral on the wire. */
export const presentationVerificationStrings = defineStrings({
  zh: {
    plan: '编辑计划',
    apply_bounded_edits: '应用限定范围的编辑',
    verify_postconditions: '验证编辑结果',
    clarify: '需要补充信息',
    verified: '已验证',
    unchanged: '无需更改',
    applied_unverified: '已应用，但未能验证',
    write_pending_quarantined: '写入可能仍在进行；已冻结后续编辑，等待收敛或重新加载',
    needs_user: '需要你确认',
    failed: '编辑失败',
    correction: '正在纠偏',
    cancelled: '已取消',
  },
  en: {
    plan: 'Edit plan',
    apply_bounded_edits: 'Apply bounded edits',
    verify_postconditions: 'Verify edit results',
    clarify: 'More information needed',
    verified: 'Verified',
    unchanged: 'No changes needed',
    applied_unverified: 'Applied, but not verified',
    write_pending_quarantined:
      'The write may still be running. Further edits are frozen; wait for reconciliation or reload.',
    needs_user: 'Your confirmation is needed',
    failed: 'Edit failed',
    correction: 'Correcting',
    cancelled: 'Cancelled',
  },
  ja: {
    plan: '編集計画',
    apply_bounded_edits: '範囲を限定して編集を適用',
    verify_postconditions: '編集結果を検証',
    clarify: '追加情報が必要です',
    verified: '検証済み',
    unchanged: '変更は不要です',
    applied_unverified: '適用済み（未検証）',
    write_pending_quarantined:
      '書き込みが続いている可能性があります。以降の編集を停止しました。収束を待つか再読み込みしてください。',
    needs_user: '確認が必要です',
    failed: '編集に失敗しました',
    correction: '修正中',
    cancelled: 'キャンセル済み',
  },
  ko: {
    plan: '편집 계획',
    apply_bounded_edits: '제한된 편집 적용',
    verify_postconditions: '편집 결과 확인',
    clarify: '추가 정보 필요',
    verified: '검증됨',
    unchanged: '변경 필요 없음',
    applied_unverified: '적용됨(검증되지 않음)',
    write_pending_quarantined:
      '쓰기가 아직 진행 중일 수 있습니다. 추가 편집이 중지되었습니다. 조정을 기다리거나 다시 로드하세요.',
    needs_user: '확인이 필요함',
    failed: '편집 실패',
    correction: '수정 중',
    cancelled: '취소됨',
  },
  fr: {
    plan: 'Plan de modification',
    apply_bounded_edits: 'Appliquer les modifications limitées',
    verify_postconditions: 'Vérifier les résultats',
    clarify: 'Informations requises',
    verified: 'Vérifié',
    unchanged: 'Aucune modification requise',
    applied_unverified: 'Appliqué, non vérifié',
    write_pending_quarantined:
      'L’écriture est peut-être toujours en cours. Les modifications suivantes sont bloquées ; attendez la réconciliation ou rechargez.',
    needs_user: 'Votre confirmation est requise',
    failed: 'Échec de la modification',
    correction: 'Correction en cours',
    cancelled: 'Annulé',
  },
  de: {
    plan: 'Bearbeitungsplan',
    apply_bounded_edits: 'Begrenzte Änderungen anwenden',
    verify_postconditions: 'Ergebnisse überprüfen',
    clarify: 'Weitere Angaben erforderlich',
    verified: 'Überprüft',
    unchanged: 'Keine Änderungen nötig',
    applied_unverified: 'Angewendet, nicht überprüft',
    write_pending_quarantined:
      'Der Schreibvorgang läuft möglicherweise noch. Weitere Änderungen sind gesperrt; warten Sie auf den Abgleich oder laden Sie neu.',
    needs_user: 'Bestätigung erforderlich',
    failed: 'Bearbeitung fehlgeschlagen',
    correction: 'Korrektur läuft',
    cancelled: 'Abgebrochen',
  },
  es: {
    plan: 'Plan de edición',
    apply_bounded_edits: 'Aplicar ediciones limitadas',
    verify_postconditions: 'Verificar los resultados',
    clarify: 'Se necesita más información',
    verified: 'Verificado',
    unchanged: 'No se requieren cambios',
    applied_unverified: 'Aplicado, sin verificar',
    write_pending_quarantined:
      'La escritura puede seguir en curso. Se bloquearon más ediciones; espera la conciliación o vuelve a cargar.',
    needs_user: 'Se necesita tu confirmación',
    failed: 'Error de edición',
    correction: 'Corrigiendo',
    cancelled: 'Cancelado',
  },
  th: {
    plan: 'แผนการแก้ไข',
    apply_bounded_edits: 'ใช้การแก้ไขในขอบเขต',
    verify_postconditions: 'ตรวจสอบผลการแก้ไข',
    clarify: 'ต้องการข้อมูลเพิ่มเติม',
    verified: 'ตรวจสอบแล้ว',
    unchanged: 'ไม่ต้องแก้ไข',
    applied_unverified: 'ใช้แล้วแต่ยังไม่ตรวจสอบ',
    write_pending_quarantined:
      'การเขียนอาจยังดำเนินอยู่ ระบบหยุดการแก้ไขเพิ่มเติมแล้ว โปรดรอการตรวจสอบหรือโหลดใหม่',
    needs_user: 'ต้องการการยืนยันจากคุณ',
    failed: 'แก้ไขไม่สำเร็จ',
    correction: 'กำลังแก้ไข',
    cancelled: 'ยกเลิกแล้ว',
  },
  id: {
    plan: 'Rencana edit',
    apply_bounded_edits: 'Terapkan edit terbatas',
    verify_postconditions: 'Verifikasi hasil edit',
    clarify: 'Perlu informasi tambahan',
    verified: 'Terverifikasi',
    unchanged: 'Tidak perlu perubahan',
    applied_unverified: 'Diterapkan, belum diverifikasi',
    write_pending_quarantined:
      'Penulisan mungkin masih berlangsung. Pengeditan lanjutan dibekukan; tunggu rekonsiliasi atau muat ulang.',
    needs_user: 'Konfirmasi Anda diperlukan',
    failed: 'Edit gagal',
    correction: 'Memperbaiki',
    cancelled: 'Dibatalkan',
  },
  ru: {
    plan: 'План изменений',
    apply_bounded_edits: 'Применить ограниченные изменения',
    verify_postconditions: 'Проверить результаты',
    clarify: 'Нужны дополнительные сведения',
    verified: 'Проверено',
    unchanged: 'Изменения не нужны',
    applied_unverified: 'Применено, но не проверено',
    write_pending_quarantined:
      'Запись может ещё выполняться. Дальнейшие изменения заблокированы; дождитесь сверки или перезагрузите документ.',
    needs_user: 'Требуется подтверждение',
    failed: 'Ошибка редактирования',
    correction: 'Исправление',
    cancelled: 'Отменено',
  },
  ar: {
    plan: 'خطة التعديل',
    apply_bounded_edits: 'تطبيق تعديلات محدودة',
    verify_postconditions: 'التحقق من النتائج',
    clarify: 'مطلوب مزيد من المعلومات',
    verified: 'تم التحقق',
    unchanged: 'لا حاجة إلى تغييرات',
    applied_unverified: 'تم التطبيق دون تحقق',
    write_pending_quarantined:
      'قد لا تزال الكتابة جارية. تم تجميد التعديلات اللاحقة؛ انتظر التسوية أو أعد التحميل.',
    needs_user: 'يلزم تأكيدك',
    failed: 'فشل التعديل',
    correction: 'جارٍ التصحيح',
    cancelled: 'تم الإلغاء',
  },
  pt: {
    plan: 'Plano de edição',
    apply_bounded_edits: 'Aplicar edições limitadas',
    verify_postconditions: 'Verificar resultados',
    clarify: 'Mais informações necessárias',
    verified: 'Verificado',
    unchanged: 'Nenhuma alteração necessária',
    applied_unverified: 'Aplicado, não verificado',
    write_pending_quarantined:
      'A gravação pode ainda estar em andamento. Novas edições foram bloqueadas; aguarde a reconciliação ou recarregue.',
    needs_user: 'Sua confirmação é necessária',
    failed: 'Falha na edição',
    correction: 'Corrigindo',
    cancelled: 'Cancelado',
  },
  it: {
    plan: 'Piano di modifica',
    apply_bounded_edits: 'Applica modifiche limitate',
    verify_postconditions: 'Verifica i risultati',
    clarify: 'Servono altre informazioni',
    verified: 'Verificato',
    unchanged: 'Nessuna modifica necessaria',
    applied_unverified: 'Applicato, non verificato',
    write_pending_quarantined:
      'La scrittura potrebbe essere ancora in corso. Le modifiche successive sono bloccate; attendi la riconciliazione o ricarica.',
    needs_user: 'È necessaria la tua conferma',
    failed: 'Modifica non riuscita',
    correction: 'Correzione in corso',
    cancelled: 'Annullato',
  },
  pl: {
    plan: 'Plan edycji',
    apply_bounded_edits: 'Zastosuj ograniczone zmiany',
    verify_postconditions: 'Zweryfikuj wyniki',
    clarify: 'Potrzebne są dodatkowe informacje',
    verified: 'Zweryfikowano',
    unchanged: 'Zmiany nie są potrzebne',
    applied_unverified: 'Zastosowano, bez weryfikacji',
    write_pending_quarantined:
      'Zapis może nadal trwać. Dalsze edycje są zablokowane; poczekaj na uzgodnienie lub przeładuj.',
    needs_user: 'Wymagane potwierdzenie',
    failed: 'Edycja nie powiodła się',
    correction: 'Korygowanie',
    cancelled: 'Anulowano',
  },
  nl: {
    plan: 'Bewerkingsplan',
    apply_bounded_edits: 'Beperkte bewerkingen toepassen',
    verify_postconditions: 'Resultaten verifiëren',
    clarify: 'Meer informatie nodig',
    verified: 'Geverifieerd',
    unchanged: 'Geen wijzigingen nodig',
    applied_unverified: 'Toegepast, niet geverifieerd',
    write_pending_quarantined:
      'De schrijfbewerking loopt mogelijk nog. Verdere bewerkingen zijn geblokkeerd; wacht op afstemming of laad opnieuw.',
    needs_user: 'Uw bevestiging is nodig',
    failed: 'Bewerken mislukt',
    correction: 'Corrigeren',
    cancelled: 'Geannuleerd',
  },
  ms: {
    plan: 'Pelan suntingan',
    apply_bounded_edits: 'Gunakan suntingan terhad',
    verify_postconditions: 'Sahkan hasil suntingan',
    clarify: 'Maklumat lanjut diperlukan',
    verified: 'Disahkan',
    unchanged: 'Tiada perubahan diperlukan',
    applied_unverified: 'Digunakan, belum disahkan',
    write_pending_quarantined:
      'Penulisan mungkin masih berjalan. Suntingan seterusnya dibekukan; tunggu penyelarasan atau muat semula.',
    needs_user: 'Pengesahan anda diperlukan',
    failed: 'Suntingan gagal',
    correction: 'Membetulkan',
    cancelled: 'Dibatalkan',
  },
  he: {
    plan: 'תוכנית עריכה',
    apply_bounded_edits: 'החלת עריכות מוגבלות',
    verify_postconditions: 'אימות תוצאות העריכה',
    clarify: 'נדרש מידע נוסף',
    verified: 'אומת',
    unchanged: 'אין צורך בשינויים',
    applied_unverified: 'הוחל, אך לא אומת',
    write_pending_quarantined:
      'ייתכן שהכתיבה עדיין מתבצעת. עריכות נוספות הוקפאו; יש להמתין להתאמה או לטעון מחדש.',
    needs_user: 'נדרש אישור ממך',
    failed: 'העריכה נכשלה',
    correction: 'מתבצע תיקון',
    cancelled: 'בוטל',
  },
  hi: {
    plan: 'संपादन योजना',
    apply_bounded_edits: 'सीमित संपादन लागू करें',
    verify_postconditions: 'संपादन परिणाम सत्यापित करें',
    clarify: 'अधिक जानकारी चाहिए',
    verified: 'सत्यापित',
    unchanged: 'बदलाव आवश्यक नहीं',
    applied_unverified: 'लागू हुआ, सत्यापित नहीं',
    write_pending_quarantined:
      'लिखना अभी जारी हो सकता है। आगे के संपादन रोक दिए गए हैं; समन्वय की प्रतीक्षा करें या पुनः लोड करें।',
    needs_user: 'आपकी पुष्टि आवश्यक है',
    failed: 'संपादन विफल',
    correction: 'सुधार जारी',
    cancelled: 'रद्द किया गया',
  },
  'zh-TW': {
    plan: '編輯計畫',
    apply_bounded_edits: '套用限定範圍的編輯',
    verify_postconditions: '驗證編輯結果',
    clarify: '需要補充資訊',
    verified: '已驗證',
    unchanged: '無需變更',
    applied_unverified: '已套用，但未能驗證',
    write_pending_quarantined: '寫入可能仍在進行；已凍結後續編輯，請等待收斂或重新載入',
    needs_user: '需要你的確認',
    failed: '編輯失敗',
    correction: '正在修正',
    cancelled: '已取消',
  },
})

export type PresentationVerificationStringKey = keyof typeof presentationVerificationStrings.zh
export function translatePresentationVerification(
  lang: Lang,
  key: PresentationVerificationStringKey,
): string {
  return presentationVerificationStrings[lang][key]
}

// ---- process-wide current language ----
// Used by Electron main-process code (shell + editor main modules share one
// bundle, so one holder). Renderers get the language over IPC instead.

let uiLang: Lang = 'zh'
const langListeners = new Set<(lang: Lang) => void>()

export function getUiLang(): Lang {
  return uiLang
}

export function setUiLang(lang: Lang): void {
  if (lang === uiLang) return
  uiLang = lang
  for (const listener of langListeners) listener(lang)
}

export function onUiLangChange(listener: (lang: Lang) => void): () => void {
  langListeners.add(listener)
  return () => langListeners.delete(listener)
}

/**
 * Build a translator over per-language dictionaries. The zh dictionary defines
 * the key set; every other language must cover exactly the same keys
 * (compile-time checked), so a missing translation is a type error, not a
 * runtime fallback.
 */
export function createI18n<D extends Record<string, string>>(dicts: LangDicts<D>) {
  return (lang: Lang, key: keyof D, params?: Params): string =>
    platformShortcuts(format(dicts[lang][key], params))
}

export type ServiceErrorCode =
  | 'auth_required'
  | 'model_credentials_missing'
  | 'model_rate_limited'
  | 'model_upstream_unavailable'
  | 'model_invalid_response'

type ServiceErrorMessages = Record<ServiceErrorCode, string>

const serviceErrors: Record<Lang, ServiceErrorMessages> = {
  zh: {
    auth_required: '请登录 WisWork 后使用 AI。',
    model_credentials_missing: 'WisWork 模型服务尚未配置。',
    model_rate_limited: 'WisWork 模型服务繁忙，请稍后重试。',
    model_upstream_unavailable: 'WisWork 模型服务暂时不可用，请稍后重试。',
    model_invalid_response: 'WisWork 模型服务返回了无效响应，请稍后重试。',
  },
  en: {
    auth_required: 'Sign in to WisWork to use AI.',
    model_credentials_missing: 'The WisWork model service is not configured.',
    model_rate_limited: 'The WisWork model service is busy. Try again shortly.',
    model_upstream_unavailable: 'The WisWork model service is temporarily unavailable.',
    model_invalid_response: 'The WisWork model service returned an invalid response.',
  },
  ja: {
    auth_required: 'AI を使用するには WisWork にサインインしてください。',
    model_credentials_missing: 'WisWork モデルサービスが設定されていません。',
    model_rate_limited:
      'WisWork モデルサービスが混み合っています。しばらくしてから再試行してください。',
    model_upstream_unavailable: 'WisWork モデルサービスは一時的に利用できません。',
    model_invalid_response: 'WisWork モデルサービスから無効な応答が返されました。',
  },
  ko: {
    auth_required: 'AI를 사용하려면 WisWork에 로그인하세요.',
    model_credentials_missing: 'WisWork 모델 서비스가 구성되지 않았습니다.',
    model_rate_limited: 'WisWork 모델 서비스가 혼잡합니다. 잠시 후 다시 시도하세요.',
    model_upstream_unavailable: 'WisWork 모델 서비스를 일시적으로 사용할 수 없습니다.',
    model_invalid_response: 'WisWork 모델 서비스가 잘못된 응답을 반환했습니다.',
  },
  fr: {
    auth_required: 'Connectez-vous à WisWork pour utiliser l’IA.',
    model_credentials_missing: 'Le service de modèles WisWork n’est pas configuré.',
    model_rate_limited: 'Le service de modèles WisWork est occupé. Réessayez bientôt.',
    model_upstream_unavailable: 'Le service de modèles WisWork est temporairement indisponible.',
    model_invalid_response: 'Le service de modèles WisWork a renvoyé une réponse invalide.',
  },
  de: {
    auth_required: 'Melden Sie sich bei WisWork an, um KI zu verwenden.',
    model_credentials_missing: 'Der WisWork-Modelldienst ist nicht konfiguriert.',
    model_rate_limited: 'Der WisWork-Modelldienst ist ausgelastet. Versuchen Sie es später erneut.',
    model_upstream_unavailable: 'Der WisWork-Modelldienst ist vorübergehend nicht verfügbar.',
    model_invalid_response: 'Der WisWork-Modelldienst hat eine ungültige Antwort geliefert.',
  },
  es: {
    auth_required: 'Inicia sesión en WisWork para usar la IA.',
    model_credentials_missing: 'El servicio de modelos de WisWork no está configurado.',
    model_rate_limited:
      'El servicio de modelos de WisWork está ocupado. Inténtalo de nuevo pronto.',
    model_upstream_unavailable:
      'El servicio de modelos de WisWork no está disponible temporalmente.',
    model_invalid_response: 'El servicio de modelos de WisWork devolvió una respuesta no válida.',
  },
  th: {
    auth_required: 'ลงชื่อเข้าใช้ WisWork เพื่อใช้ AI',
    model_credentials_missing: 'ยังไม่ได้กำหนดค่าบริการโมเดล WisWork',
    model_rate_limited: 'บริการโมเดล WisWork ไม่ว่าง โปรดลองอีกครั้งภายหลัง',
    model_upstream_unavailable: 'บริการโมเดล WisWork ไม่พร้อมใช้งานชั่วคราว',
    model_invalid_response: 'บริการโมเดล WisWork ส่งการตอบกลับที่ไม่ถูกต้อง',
  },
  id: {
    auth_required: 'Masuk ke WisWork untuk menggunakan AI.',
    model_credentials_missing: 'Layanan model WisWork belum dikonfigurasi.',
    model_rate_limited: 'Layanan model WisWork sedang sibuk. Coba lagi nanti.',
    model_upstream_unavailable: 'Layanan model WisWork sementara tidak tersedia.',
    model_invalid_response: 'Layanan model WisWork memberikan respons yang tidak valid.',
  },
  ru: {
    auth_required: 'Войдите в WisWork, чтобы использовать ИИ.',
    model_credentials_missing: 'Служба моделей WisWork не настроена.',
    model_rate_limited: 'Служба моделей WisWork занята. Повторите попытку позже.',
    model_upstream_unavailable: 'Служба моделей WisWork временно недоступна.',
    model_invalid_response: 'Служба моделей WisWork вернула недопустимый ответ.',
  },
  ar: {
    auth_required: 'سجّل الدخول إلى WisWork لاستخدام الذكاء الاصطناعي.',
    model_credentials_missing: 'لم تتم تهيئة خدمة نماذج WisWork.',
    model_rate_limited: 'خدمة نماذج WisWork مشغولة. حاول مجددًا بعد قليل.',
    model_upstream_unavailable: 'خدمة نماذج WisWork غير متاحة مؤقتًا.',
    model_invalid_response: 'أعادت خدمة نماذج WisWork استجابة غير صالحة.',
  },
  pt: {
    auth_required: 'Entre no WisWork para usar a IA.',
    model_credentials_missing: 'O serviço de modelos do WisWork não está configurado.',
    model_rate_limited: 'O serviço de modelos do WisWork está ocupado. Tente novamente em breve.',
    model_upstream_unavailable:
      'O serviço de modelos do WisWork está temporariamente indisponível.',
    model_invalid_response: 'O serviço de modelos do WisWork retornou uma resposta inválida.',
  },
  it: {
    auth_required: 'Accedi a WisWork per usare l’IA.',
    model_credentials_missing: 'Il servizio modelli WisWork non è configurato.',
    model_rate_limited: 'Il servizio modelli WisWork è occupato. Riprova tra poco.',
    model_upstream_unavailable: 'Il servizio modelli WisWork è temporaneamente non disponibile.',
    model_invalid_response: 'Il servizio modelli WisWork ha restituito una risposta non valida.',
  },
  pl: {
    auth_required: 'Zaloguj się do WisWork, aby korzystać z AI.',
    model_credentials_missing: 'Usługa modeli WisWork nie jest skonfigurowana.',
    model_rate_limited: 'Usługa modeli WisWork jest zajęta. Spróbuj ponownie później.',
    model_upstream_unavailable: 'Usługa modeli WisWork jest tymczasowo niedostępna.',
    model_invalid_response: 'Usługa modeli WisWork zwróciła nieprawidłową odpowiedź.',
  },
  nl: {
    auth_required: 'Meld u aan bij WisWork om AI te gebruiken.',
    model_credentials_missing: 'De WisWork-modelservice is niet geconfigureerd.',
    model_rate_limited: 'De WisWork-modelservice is bezet. Probeer het later opnieuw.',
    model_upstream_unavailable: 'De WisWork-modelservice is tijdelijk niet beschikbaar.',
    model_invalid_response: 'De WisWork-modelservice heeft een ongeldig antwoord gegeven.',
  },
  ms: {
    auth_required: 'Log masuk ke WisWork untuk menggunakan AI.',
    model_credentials_missing: 'Perkhidmatan model WisWork belum dikonfigurasi.',
    model_rate_limited: 'Perkhidmatan model WisWork sedang sibuk. Cuba lagi sebentar lagi.',
    model_upstream_unavailable: 'Perkhidmatan model WisWork tidak tersedia buat sementara waktu.',
    model_invalid_response: 'Perkhidmatan model WisWork mengembalikan respons yang tidak sah.',
  },
  he: {
    auth_required: 'יש להתחבר ל-WisWork כדי להשתמש ב-AI.',
    model_credentials_missing: 'שירות המודלים של WisWork אינו מוגדר.',
    model_rate_limited: 'שירות המודלים של WisWork עמוס. נסו שוב מאוחר יותר.',
    model_upstream_unavailable: 'שירות המודלים של WisWork אינו זמין זמנית.',
    model_invalid_response: 'שירות המודלים של WisWork החזיר תגובה לא תקינה.',
  },
  hi: {
    auth_required: 'AI का उपयोग करने के लिए WisWork में साइन इन करें।',
    model_credentials_missing: 'WisWork मॉडल सेवा कॉन्फ़िगर नहीं है।',
    model_rate_limited: 'WisWork मॉडल सेवा व्यस्त है। थोड़ी देर बाद फिर प्रयास करें।',
    model_upstream_unavailable: 'WisWork मॉडल सेवा अस्थायी रूप से उपलब्ध नहीं है।',
    model_invalid_response: 'WisWork मॉडल सेवा ने अमान्य प्रतिक्रिया लौटाई।',
  },
  'zh-TW': {
    auth_required: '請登入 WisWork 後使用 AI。',
    model_credentials_missing: 'WisWork 模型服務尚未設定。',
    model_rate_limited: 'WisWork 模型服務忙碌中，請稍後再試。',
    model_upstream_unavailable: 'WisWork 模型服務暫時無法使用。',
    model_invalid_response: 'WisWork 模型服務傳回無效回應。',
  },
}

export function translateServiceError(lang: Lang, code: ServiceErrorCode): string {
  return serviceErrors[lang][code]
}
