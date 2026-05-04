# -*- coding: utf-8 -*-
"""
Система проверки приказов Министерства экономического развития России
Поддержка форматов: TXT, DOCX, PDF
"""

import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
import os
import threading
import webbrowser
import csv
from datetime import datetime
from order_structure_validator import OrderStructureValidator
from data_loader import load_file_by_type
from error_detection_model import ErrorDetectionTrainer
import config


SECTIONS = [
    {
        "title": "Раздел 1",
        "full_title": "Проекты приказов МЭР и результаты общественного обсуждения",
        "url": "https://regulation.gov.ru/projects/?type=Grid#departments=6",
        "portal": "regulation.gov.ru",
        "columns": ("№", "Название проекта приказа", "Дата публикации", "Статус", "Комментариев"),
        "data": [
            ("1", "Об утверждении методических рекомендаций по оценке регулирующего воздействия", "15.03.2025", "На обсуждении", "12"),
            ("2", "О внесении изменений в приказ МЭР № 532 от 25.09.2024", "10.03.2025", "Обсуждение завершено", "8"),
            ("3", "Об установлении порядка ведения реестра инвестиционных проектов", "05.03.2025", "Обсуждение завершено", "15"),
            ("4", "Об утверждении формы заключения об ОРВ проектов нормативных актов", "28.02.2025", "Обсуждение завершено", "6"),
            ("5", "О внесении изменений в порядок проведения антикоррупционной экспертизы", "20.02.2025", "На доработке", "22"),
            ("6", "Об утверждении стандарта государственной услуги по регистрации юрлиц", "15.02.2025", "Обсуждение завершено", "9"),
            ("7", "О признании утратившими силу некоторых приказов МЭР", "10.02.2025", "Обсуждение завершено", "3"),
            ("8", "Об установлении критериев оценки эффективности деятельности органов власти", "05.02.2025", "На обсуждении", "17"),
        ],
    },
    {
        "title": "Раздел 2",
        "full_title": "Независимая антикоррупционная экспертиза проектов приказов",
        "url": "https://regulation.gov.ru/projects/?type=Grid#departments=6&categories=5",
        "portal": "regulation.gov.ru",
        "columns": ("№", "Проект нормативного акта", "Дата начала экспертизы", "Эксперт", "Результат"),
        "data": [
            ("1", "Проект приказа об утверждении порядка ведения реестра субъектов МСП", "20.03.2025", "Иванов А.В.", "Нарушений не выявлено"),
            ("2", "Проект приказа о внесении изменений в методику оценки ущерба", "18.03.2025", "Петрова Е.С.", "Выявлены коррупциогенные факторы"),
            ("3", "Проект приказа об установлении требований к раскрытию информации", "12.03.2025", "Сидоров К.М.", "Нарушений не выявлено"),
            ("4", "Проект приказа об утверждении административного регламента", "08.03.2025", "Козлова Н.И.", "Нарушений не выявлено"),
            ("5", "Проект приказа о порядке согласования инвестиционных программ", "02.03.2025", "Новиков Д.А.", "На рассмотрении"),
            ("6", "Проект приказа об утверждении перечня документов для господдержки МСП", "25.02.2025", "Орлов В.Г.", "Нарушений не выявлено"),
            ("7", "Проект приказа о признании утратившими силу ряда нормативных актов", "20.02.2025", "Федоров С.Р.", "Нарушений не выявлено"),
        ],
    },
    {
        "title": "Раздел 3",
        "full_title": "Участие в подготовке законодательных актов. Сводная таблица по субъектам ЗИ с итогами",
        "url": "https://sozd.duma.gov.ru/oz_info_spzi/spzi_list",
        "portal": "sozd.duma.gov.ru",
        "columns": ("№", "Субъект законодательной инициативы", "Законопроект", "Стадия", "Дата внесения"),
        "data": [
            ("1", "Правительство РФ", "О внесении изменений в ФЗ «О развитии малого и среднего предпринимательства»", "Второе чтение", "22.03.2025"),
            ("2", "МЭР России", "О внесении изменений в ФЗ «Об обществах с ограниченной ответственностью»", "Первое чтение", "15.03.2025"),
            ("3", "Депутаты ГД", "Об инновационной деятельности и государственной инновационной политике", "Внесен", "10.03.2025"),
            ("4", "Правительство РФ", "О внесении изменений в Федеральный закон «Об инвестиционной деятельности»", "Подписан", "05.03.2025"),
            ("5", "Совет Федерации", "О внесении изменений в ФЗ «О государственной регистрации юридических лиц»", "Третье чтение", "28.02.2025"),
            ("6", "МЭР России", "О внесении изменений в Налоговый кодекс РФ (в части налогообложения МСП)", "Первое чтение", "20.02.2025"),
            ("7", "Правительство РФ", "О внесении изменений в ФЗ «О несостоятельности (банкротстве)»", "На рассмотрении", "14.02.2025"),
            ("8", "Депутаты ГД", "Об особых экономических зонах в Российской Федерации (новая редакция)", "Внесен", "08.02.2025"),
        ],
    },
    {
        "title": "Раздел 4",
        "full_title": "Официально опубликованные нормативные акты МЭР",
        "url": "http://publication.pravo.gov.ru/search/federal_authorities?pageSize=30&index=1&SignatoryAuthorityId=59ad7c35-7d14-424e-b305-f964615734bd&&PublishDateSearchType=0&NumberSearchType=0&DocumentDateSearchType=0&JdRegSearchType=0&SortedBy=6&SortDestination=1",
        "portal": "publication.pravo.gov.ru",
        "columns": ("№", "Название нормативного акта", "Дата публикации", "Номер приказа", "Статус"),
        "data": [
            ("1", "Приказ МЭР об утверждении методики расчёта прогнозного плана приватизации", "25.03.2025", "№ 145", "Действующий"),
            ("2", "Приказ МЭР об утверждении методики оценки финансового состояния юрлиц", "20.03.2025", "№ 142", "Действующий"),
            ("3", "Приказ МЭР о внесении изменений в приказ № 112 от 14.01.2025", "15.03.2025", "№ 138", "Действующий"),
            ("4", "Приказ МЭР об утверждении административного регламента предоставления госуслуги", "10.03.2025", "№ 134", "Действующий"),
            ("5", "Приказ МЭР о порядке согласования государственных программ субъектов РФ", "05.03.2025", "№ 129", "Действующий"),
            ("6", "Приказ МЭР об установлении требований к структуре стратегий развития", "28.02.2025", "№ 125", "Действующий"),
            ("7", "Приказ МЭР о признании утратившим силу приказа № 98 от 22.11.2024", "20.02.2025", "№ 118", "Действующий"),
            ("8", "Приказ МЭР об утверждении перечня системообразующих организаций", "15.02.2025", "№ 114", "Действующий"),
            ("9", "Приказ МЭР о порядке ведения реестра контрактов по ГЧП", "10.02.2025", "№ 110", "Действующий"),
        ],
    },
]

# Справочник и шаблоны для антикоррупционного сканирования формулировок НПА
CORRUPTION_REFERENCE_TEXT = """
При проведении анализа нормативных правовых актов к коррупциогенным формулировкам
целесообразно относить следующие конструкции:

  «Орган может принять иные решения по своему усмотрению»
      (отсутствуют пределы усмотрения и критерии выбора решения)

  «Допускается отклонение от установленных требований»
      (не указаны условия и границы допустимого отклонения)

  «При необходимости заявитель предоставляет дополнительные документы»
      (не определено, в каких случаях возникает необходимость и какие документы требуются)

  «Компетентный орган вправе установить особый порядок»
      (не раскрыто содержание «особого порядка»)

  «Решение принимается в разумный срок»
      (оценочная категория без конкретных временных рамок)

  «В исключительных случаях допускается продление срока»
      (не определены критерии «исключительности»)

  «Иные основания, предусмотренные действующим законодательством»
      (слишком широкая отсылка без конкретизации)

  «При наличии обоснованных причин может быть отказано»
      (не раскрыт перечень или критерии «обоснованных причин»)

  «Орган вправе запросить дополнительные пояснения»
      (не указано, какие именно пояснения и в каком объёме)

  «Размер выплаты определяется индивидуально»
      (отсутствуют методика расчёта и объективные критерии)

  «Допускается применение иных мер воздействия»
      (не конкретизирован перечень мер)

  «Решение принимается с учетом специфики ситуации»
      (размытая формулировка без перечня факторов)

  «Уполномоченное лицо самостоятельно определяет порядок действий»
      (чрезмерная дискреция без регламентации)

  «В отдельных случаях требования могут не применяться»
      (не раскрыто, какие случаи считаются отдельными)

  «Орган может учитывать иные обстоятельства»
      (отсутствует исчерпывающий перечень обстоятельств)
""".strip()

# (краткое название, regex, пояснение риска) — гибкий поиск по тексту акта
CORRUPTION_SCAN_RULES = [
    (
        "Орган может принять иные решения по своему усмотрению",
        r"орган\w{0,24}\s+(?:может\s+)?(?:принять\s+)?иные\s+решени\w*\s+.{0,45}усмотрени\w*",
        "отсутствуют пределы усмотрения и критерии выбора решения",
    ),
    (
        "Допускается отклонение от установленных требований",
        r"допускается\s+отклонение.{0,55}установленн\w*\s+требовани\w*",
        "не указаны условия и границы допустимого отклонения",
    ),
    (
        "При необходимости заявитель предоставляет дополнительные документы",
        r"при\s+необходимости.{0,40}заявител\w*.{0,50}дополнительн\w*\s+документ",
        "не определено, в каких случаях возникает необходимость и какие документы требуются",
    ),
    (
        "Компетентный орган вправе установить особый порядок",
        r"компетентн\w*\s+орган\w*.{0,55}особ\w+\s+порядок",
        "не раскрыто содержание «особого порядка»",
    ),
    (
        "Решение принимается в разумный срок",
        r"решени\w*\s+.{0,18}принима\w*.{0,30}разумн\w*\s+срок",
        "оценочная категория без конкретных временных рамок",
    ),
    (
        "В исключительных случаях допускается продление срока",
        r"исключительн\w*\s+случа\w*.{0,45}продлени\w*\s+срок",
        "не определены критерии «исключительности»",
    ),
    (
        "Иные основания, предусмотренные действующим законодательством",
        r"иные\s+основани\w*.{0,65}законодательств",
        "слишком широкая отсылка без конкретизации",
    ),
    (
        "При наличии обоснованных причин может быть отказано",
        r"обоснованн\w*\s+причин.{0,45}(?:может\s+быть\s+)?отказан",
        "не раскрыт перечень или критерии «обоснованных причин»",
    ),
    (
        "Орган вправе запросить дополнительные пояснения",
        r"орган\w{0,20}\s+вправе\s+запросить.{0,30}дополнительн\w*\s+пояснени",
        "не указано, какие именно пояснения и в каком объёме",
    ),
    (
        "Размер выплаты определяется индивидуально",
        r"размер\s+выплат\w*.{0,35}индивидуально",
        "отсутствуют методика расчёта и объективные критерии",
    ),
    (
        "Допускается применение иных мер воздействия",
        r"допускается\s+применени\w*.{0,30}иных\s+мер\s+воздействи",
        "не конкретизирован перечень мер",
    ),
    (
        "Решение принимается с учетом специфики ситуации",
        r"решени\w*\s+.{0,22}уч[её]том\s+специфики.{0,30}ситуаци",
        "размытая формулировка без перечня факторов",
    ),
    (
        "Уполномоченное лицо самостоятельно определяет порядок действий",
        r"уполномоченн\w*\s+лиц\w*.{0,45}самостоятельно\s+определ\w*.{0,45}порядок\s+действи",
        "чрезмерная дискреция без регламентации",
    ),
    (
        "В отдельных случаях требования могут не применяться",
        r"отдельн\w*\s+случа\w*.{0,50}требовани\w*\s+могут\s+не\s+применяться",
        "не раскрыто, какие случаи считаются отдельными",
    ),
    (
        "Орган может учитывать иные обстоятельства",
        r"орган\w{0,20}\s+может\s+учитывать.{0,28}иные\s+обстоятельств",
        "отсутствует исчерпывающий перечень обстоятельств",
    ),
]


class OrderCheckerGUI:
    BLUE = "#0039A6"
    RED = "#D52B1E"
    WHITE = "#FFFFFF"
    GREEN = "#2ECC71"
    LIGHT_BLUE = "#3498DB"
    DARK_BLUE = "#1A5490"
    GRAY = "#F0F2F5"
    SIDEBAR_BG = "#1E3A5F"
    SIDEBAR_ACTIVE = "#0039A6"
    SIDEBAR_TEXT = "#FFFFFF"

    def __init__(self, root):
        self.root = root
        self.root.title("Система проверки приказов Минэкономразвития России")
        self.root.geometry("1280x780")
        self.root.minsize(1024, 660)
        self.root.configure(bg=self.GRAY)

        self.validator = OrderStructureValidator()
        self.model_available = False
        self.trainer = None
        self._load_model()

        self.current_file = None
        self.current_text = None
        self.last_validation_result = None
        self.last_model_result = None
        self.active_section = None
        self.current_tree = None

        self._build_ui()

    # ------------------------------------------------------------------ model
    def _load_model(self):
        try:
            model_path = os.path.join(config.MODELS_DIR, "best_error_detection_model.pt")
            if os.path.exists(model_path):
                self.trainer = ErrorDetectionTrainer()
                self.trainer.load_model("best_error_detection_model.pt")
                self.model_available = True
                print("✓ Модель загружена")
            else:
                print("⚠ Модель не найдена, используется только проверка структуры")
        except Exception as e:
            print(f"Ошибка загрузки модели: {e}")

    # ------------------------------------------------------------------ build
    def _build_ui(self):
        self._build_header()
        self._build_toolbar()
        self._build_file_info()
        self._build_main()
        self._build_footer()

    def _build_header(self):
        hf = tk.Frame(self.root, bg=self.RED, height=58)
        hf.pack(fill=tk.X)
        hf.pack_propagate(False)

        tk.Label(
            hf,
            text="СИСТЕМА ПРОВЕРКИ ПРИКАЗОВ МИНЭКОНОМРАЗВИТИЯ РОССИИ",
            font=("Arial", 13, "bold"),
            bg=self.RED, fg=self.WHITE,
        ).pack(side=tk.LEFT, padx=20, pady=10)

        status_text = "✓ Модель активна" if self.model_available else "⚠ Только структура"
        status_bg = "#1E8449" if self.model_available else "#E67E22"
        tk.Label(
            hf, text=status_text,
            font=("Arial", 9, "bold"),
            bg=status_bg, fg=self.WHITE,
            padx=10, pady=4, relief=tk.RAISED,
        ).pack(side=tk.RIGHT, padx=15, pady=12)

    def _build_toolbar(self):
        tf = tk.Frame(self.root, bg=self.BLUE, padx=15, pady=10)
        tf.pack(fill=tk.X)

        btn_cfg = dict(font=("Arial", 10, "bold"), padx=14, pady=7,
                       cursor="hand2", relief=tk.RAISED, bd=2)

        self.load_btn = tk.Button(
            tf, text="📂 Загрузить документ",
            command=self.load_file,
            bg=self.WHITE, fg=self.BLUE, **btn_cfg)
        self.load_btn.pack(side=tk.LEFT, padx=4)

        self.check_btn = tk.Button(
            tf, text="✓ Проверить структуру",
            command=self.check_document,
            bg="#1E8449", fg=self.WHITE,
            state=tk.DISABLED, **btn_cfg)
        self.check_btn.pack(side=tk.LEFT, padx=4)

        self.save_btn = tk.Button(
            tf, text="💾 Скачать отчёт",
            command=self.save_report,
            bg=self.LIGHT_BLUE, fg=self.WHITE,
            state=tk.DISABLED, **btn_cfg)
        self.save_btn.pack(side=tk.LEFT, padx=4)

        self.clear_btn = tk.Button(
            tf, text="🗑 Очистить",
            command=self.clear_results,
            bg=self.RED, fg=self.WHITE, **btn_cfg)
        self.clear_btn.pack(side=tk.LEFT, padx=4)

    def _build_file_info(self):
        self.info_bar = tk.Frame(self.root, bg=self.WHITE,
                                 padx=12, pady=6, relief=tk.SOLID, bd=1)
        self.info_bar.pack(fill=tk.X, padx=8, pady=(4, 0))

        self.file_label = tk.Label(
            self.info_bar, text="📄 Файл не загружен",
            font=("Arial", 10), bg=self.WHITE, fg=self.BLUE, anchor="w")
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("G.Horizontal.TProgressbar",
                        background=self.GREEN, troughcolor=self.WHITE,
                        bordercolor=self.BLUE, lightcolor=self.GREEN, darkcolor=self.GREEN)
        self.progress = ttk.Progressbar(
            self.info_bar, mode="indeterminate",
            length=200, style="G.Horizontal.TProgressbar")

    def _build_main(self):
        main = tk.Frame(self.root, bg=self.GRAY)
        main.pack(fill=tk.BOTH, expand=True, padx=8, pady=5)

        self._build_sidebar(main)
        self._build_content(main)

    def _build_sidebar(self, parent):
        sb = tk.Frame(parent, bg=self.SIDEBAR_BG, width=235)
        sb.pack(side=tk.LEFT, fill=tk.Y)
        sb.pack_propagate(False)

        tk.Label(
            sb, text="РАЗДЕЛЫ",
            font=("Arial", 10, "bold"),
            bg=self.SIDEBAR_BG, fg=self.WHITE, pady=12,
        ).pack(fill=tk.X)

        tk.Frame(sb, bg="#4A90D9", height=1).pack(fill=tk.X, padx=8)

        self._section_btns = []
        for idx, sec in enumerate(SECTIONS):
            short = sec["full_title"]
            if len(short) > 42:
                short = short[:42] + "…"
            btn = tk.Button(
                sb,
                text=f"{sec['title']}\n{short}",
                font=("Arial", 8),
                bg=self.SIDEBAR_BG, fg=self.WHITE,
                activebackground=self.SIDEBAR_ACTIVE, activeforeground=self.WHITE,
                cursor="hand2", relief=tk.FLAT, bd=0,
                wraplength=215, justify=tk.LEFT, anchor="w",
                padx=10, pady=9,
                command=lambda i=idx: self._switch_section(i),
            )
            btn.pack(fill=tk.X, padx=2, pady=1)
            self._section_btns.append(btn)

        tk.Frame(sb, bg="#4A90D9", height=1).pack(fill=tk.X, padx=8, pady=(12, 6))

        self.portal_btn = tk.Button(
            sb, text="🌐 Открыть на портале",
            command=self._open_portal,
            font=("Arial", 9, "bold"),
            bg="#27AE60", fg=self.WHITE,
            cursor="hand2", relief=tk.RAISED, bd=2,
            padx=10, pady=8, state=tk.DISABLED,
        )
        self.portal_btn.pack(fill=tk.X, padx=10, pady=3)

        self.export_btn = tk.Button(
            sb, text="📊 Экспорт в CSV",
            command=self._export_csv,
            font=("Arial", 9, "bold"),
            bg="#8E44AD", fg=self.WHITE,
            cursor="hand2", relief=tk.RAISED, bd=2,
            padx=10, pady=8, state=tk.DISABLED,
        )
        self.export_btn.pack(fill=tk.X, padx=10, pady=3)

    def _build_content(self, parent):
        cf = tk.Frame(parent, bg=self.WHITE, relief=tk.SOLID, bd=1)
        cf.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))

        # content header bar
        ch = tk.Frame(cf, bg=self.DARK_BLUE, height=36)
        ch.pack(fill=tk.X)
        ch.pack_propagate(False)
        self.content_title = tk.Label(
            ch, text="📊 РЕЗУЛЬТАТЫ ПРОВЕРКИ",
            font=("Arial", 11, "bold"),
            bg=self.DARK_BLUE, fg=self.WHITE,
            anchor="w", padx=12,
        )
        self.content_title.pack(fill=tk.BOTH, expand=True)

        # doc results pane
        self.doc_pane = tk.Frame(cf, bg=self.WHITE)
        self.doc_pane.pack(fill=tk.BOTH, expand=True)

        self.result_text = scrolledtext.ScrolledText(
            self.doc_pane,
            font=("Consolas", 10), wrap=tk.WORD,
            bg=self.WHITE, relief=tk.FLAT, bd=0,
            padx=15, pady=15,
        )
        self.result_text.pack(fill=tk.BOTH, expand=True)

        self.result_text.tag_config("success",  foreground=self.GREEN,     font=("Consolas", 11, "bold"))
        self.result_text.tag_config("error",    foreground=self.RED,       font=("Consolas", 11, "bold"))
        self.result_text.tag_config("warning",  foreground="#F39C12",      font=("Consolas", 10, "bold"))
        self.result_text.tag_config("header",   foreground=self.BLUE,      font=("Consolas", 12, "bold"))
        self.result_text.tag_config("section",  foreground=self.DARK_BLUE, font=("Consolas", 10, "bold"))
        self.result_text.tag_config("normal",   foreground="#2C3E50",      font=("Consolas", 10))
        self.result_text.tag_config("corr",     foreground="#C0392B",      font=("Consolas", 10, "bold"))
        self.result_text.tag_config("corr_note", foreground="#566573",    font=("Consolas", 9))

        # section data pane (hidden initially)
        self.sec_pane = tk.Frame(cf, bg=self.WHITE)

    def _build_footer(self):
        ff = tk.Frame(self.root, bg=self.BLUE, height=30)
        ff.pack(fill=tk.X, side=tk.BOTTOM)
        ff.pack_propagate(False)
        tk.Label(
            ff,
            text="Поддержка форматов: TXT • DOCX • PDF  |  © Министерство экономического развития России 2025",
            font=("Arial", 8), bg=self.BLUE, fg=self.WHITE,
        ).pack(pady=5)

    # ---------------------------------------------------------------- sidebar
    def _switch_section(self, idx):
        self.active_section = idx
        sec = SECTIONS[idx]

        for i, btn in enumerate(self._section_btns):
            btn.config(bg=self.SIDEBAR_ACTIVE if i == idx else self.SIDEBAR_BG,
                       relief=tk.RIDGE if i == idx else tk.FLAT)

        self.portal_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.NORMAL)

        self.doc_pane.pack_forget()
        self.sec_pane.pack(fill=tk.BOTH, expand=True)

        for w in self.sec_pane.winfo_children():
            w.destroy()

        # section info bar
        info = tk.Frame(self.sec_pane, bg="#E8EFF7", pady=8)
        info.pack(fill=tk.X)
        tk.Label(info, text=sec["full_title"],
                 font=("Arial", 10, "bold"), bg="#E8EFF7", fg=self.DARK_BLUE,
                 wraplength=800, justify=tk.LEFT, anchor="w", padx=12).pack(fill=tk.X)
        tk.Label(info, text=f"Источник данных: {sec['portal']}",
                 font=("Arial", 9), bg="#E8EFF7", fg="#555555",
                 anchor="w", padx=12).pack(fill=tk.X)

        # table
        tbl_frame = tk.Frame(self.sec_pane, bg=self.WHITE)
        tbl_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        style = ttk.Style()
        style.configure("Tbl.Treeview", rowheight=26, font=("Arial", 9))
        style.configure("Tbl.Treeview.Heading",
                        font=("Arial", 9, "bold"),
                        background=self.BLUE, foreground="white")
        style.map("Tbl.Treeview.Heading", background=[("active", self.DARK_BLUE)])

        cols = sec["columns"]
        tree = ttk.Treeview(tbl_frame, columns=cols, show="headings",
                            height=14, style="Tbl.Treeview")

        col_widths = {"№": 38, "Дата публикации": 120, "Дата начала экспертизы": 160,
                      "Дата внесения": 120, "Статус": 160, "Результат": 200,
                      "Стадия": 140, "Номер приказа": 110, "Комментариев": 110,
                      "Экспертов": 80}
        for col in cols:
            w = col_widths.get(col, 240)
            anchor = "center" if col in ("№", "Комментариев", "Экспертов", "Номер приказа") else "w"
            tree.heading(col, text=col)
            tree.column(col, width=w, anchor=anchor, minwidth=40)

        for i, row in enumerate(sec["data"]):
            tree.insert("", tk.END, values=row, tags=("even" if i % 2 == 0 else "odd",))

        tree.tag_configure("even", background="#F4F6FB")
        tree.tag_configure("odd",  background=self.WHITE)

        vsb = ttk.Scrollbar(tbl_frame, orient="vertical",   command=tree.yview)
        hsb = ttk.Scrollbar(tbl_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT,  fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        tree.pack(fill=tk.BOTH, expand=True)

        self.current_tree = tree

        title_short = sec["full_title"]
        if len(title_short) > 55:
            title_short = title_short[:55] + "…"
        self.content_title.config(text=f"📋 {sec['title'].upper()}: {title_short}")

    def _show_doc_pane(self):
        self.sec_pane.pack_forget()
        self.doc_pane.pack(fill=tk.BOTH, expand=True)
        self.content_title.config(text="📊 РЕЗУЛЬТАТЫ ПРОВЕРКИ")
        for btn in self._section_btns:
            btn.config(bg=self.SIDEBAR_BG, relief=tk.FLAT)
        self.active_section = None
        self.portal_btn.config(state=tk.DISABLED)
        self.export_btn.config(state=tk.DISABLED)

    def _open_portal(self):
        if self.active_section is not None:
            webbrowser.open(SECTIONS[self.active_section]["url"])

    def _export_csv(self):
        if self.active_section is None:
            messagebox.showwarning("Предупреждение", "Сначала выберите раздел")
            return

        sec = SECTIONS[self.active_section]
        path = filedialog.asksaveasfilename(
            title="Экспорт в CSV",
            defaultextension=".csv",
            filetypes=[("CSV файл", "*.csv"), ("Все файлы", "*.*")],
            initialfile=f"раздел{self.active_section + 1}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        )
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(sec["columns"])
                w.writerows(sec["data"])
            messagebox.showinfo("Экспорт", f"Данные сохранены:\n{os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при экспорте:\n{e}")

    # --------------------------------------------------------- helpers / scan
    @staticmethod
    def _doc_stats(text: str) -> dict:
        """Базовая статистика документа."""
        import re
        lines  = text.splitlines()
        words  = re.findall(r'\w+', text)
        paras  = [p.strip() for p in re.split(r'\n{2,}', text) if p.strip()]
        sents  = re.split(r'[.!?]+', text)
        sents  = [s.strip() for s in sents if len(s.strip()) > 10]
        return {
            "chars":   len(text),
            "lines":   len(lines),
            "words":   len(words),
            "paras":   len(paras),
            "sents":   len(sents),
        }

    @staticmethod
    def _quick_scan(text: str) -> list:
        """
        Быстрая проверка присутствия ключевых элементов приказа.
        Возвращает список (название, найдено:bool, найденный фрагмент).
        """
        import re
        checks = []

        def first_match(patterns, t, flags=re.IGNORECASE):
            for p in patterns:
                m = re.search(p, t, flags)
                if m:
                    return m.group(0)[:60].strip()
            return None

        head = "\n".join(text.splitlines()[:25])

        # 1. Наименование органа
        frag = first_match([r'МИНИСТЕРСТВО[^\n]+', r'Минэкономразвития[^\n]+',
                             r'ФЕДЕРАЛЬН[А-ЯЁ ]+[^\n]+'], head)
        checks.append(("Наименование органа (шапка)", frag is not None, frag or ""))

        # 2. Слово ПРИКАЗ
        frag = first_match([r'^\s*ПРИКАЗ\s*$'], head, re.MULTILINE | re.IGNORECASE)
        if not frag:
            frag = first_match([r'ПРИКАЗ'], head)
        checks.append(("Слово «ПРИКАЗ»", frag is not None, frag or ""))

        # 3. Дата
        frag = first_match([r'«?\d{1,2}»?\s+[а-яё]+\s+\d{4}',
                             r'\d{2}\.\d{2}\.\d{4}'], text)
        checks.append(("Дата издания", frag is not None, frag or ""))

        # 4. Номер
        frag = first_match([r'№\s*[\d\w/-]+'], text)
        checks.append(("Регистрационный номер (№)", frag is not None, frag or ""))

        # 5. Место издания
        frag = first_match([r'г\.\s*[А-ЯЁ][а-яё-]+'], head)
        checks.append(("Место издания", frag is not None, frag or ""))

        # 6. Заголовок «О …»
        frag = first_match([r'^\s*О[бб]?\s+[А-ЯЁа-яё].{10,}'], text, re.MULTILINE)
        checks.append(("Заголовок к тексту (О / Об …)", frag is not None, frag or ""))

        # 7. Преамбула
        frag = first_match([r'В\s+соответствии\s+с\s+.{5,}',
                             r'На\s+основании\s+.{5,}',
                             r'Во\s+исполнение\s+.{5,}',
                             r'Руководствуясь\s+.{5,}'], text)
        checks.append(("Преамбула (основание издания)", frag is not None, frag or ""))

        # 8. Ссылка на НПА
        frag = first_match([r'[Фф]едеральн\w+ закон\w*[^\n]{0,40}',
                             r'[Пп]остановлени\w+ Правительства[^\n]{0,40}',
                             r'№\s*\d+[-\w]*\s+от\s+\d{2}\.\d{2}\.\d{4}'], text)
        checks.append(("Ссылка на НПА в преамбуле", frag is not None, frag or ""))

        # 9. ПРИКАЗЫВАЮ
        frag = first_match([r'ПРИКАЗЫВАЮ\s*:?'], text)
        checks.append(("Распорядительное слово «ПРИКАЗЫВАЮ:»", frag is not None, frag or ""))

        # 10. Нумерованные пункты
        items = re.findall(r'(?:^|\n)\s*\d+\.\s+\S', text, re.MULTILINE)
        frag = f"пунктов найдено: {len(items)}" if items else ""
        checks.append(("Нумерованные пункты в распорядительной части",
                        len(items) > 0, frag))

        # 11. Контроль за исполнением
        frag = first_match([r'[Кк]онтроль\s+за\s+исполнением[^\n]{0,50}',
                             r'[Кк]онтроль\s+исполнения[^\n]{0,50}'], text)
        checks.append(("Пункт о контроле исполнения", frag is not None, frag or ""))

        # 12. Подпись / должность подписанта
        frag = first_match([r'[Мм]инистр\s+.{0,40}',
                             r'[Зз]аместитель\s+министра\s+.{0,40}',
                             r'[Рр]уководитель\s+.{0,40}'], text)
        checks.append(("Подпись (должность подписанта)", frag is not None, frag or ""))

        # 13. Приложения
        frag = first_match([r'[Пп]риложение\s*(?:№\s*)?\d*[^\n]{0,40}'], text)
        checks.append(("Приложения", frag is not None, frag or "нет ссылок на приложения"))

        # 14. Регистрация в Минюсте (информационно)
        frag = first_match([r'[Зз]арегистрировано\s+в\s+[Мм]инюст[^\n]{0,50}'], text)
        checks.append(("Зарегистрировано в Минюсте", frag is not None, frag or "нет"))

        return checks

    @staticmethod
    def _scan_corruption_risks(text: str):
        """
        Ищет в тексте НПА фрагменты, близкие к типовым коррупциогенным конструкциям.
        Возвращает список словарей: label, risk, snippet.
        """
        import re
        hits = []
        seen_labels = set()
        flags = re.IGNORECASE | re.DOTALL
        for label, pat, risk in CORRUPTION_SCAN_RULES:
            if label in seen_labels:
                continue
            m = re.search(pat, text, flags)
            if not m:
                continue
            seen_labels.add(label)
            a, b = m.span()
            lo = max(0, a - 50)
            hi = min(len(text), b + 70)
            snippet = text[lo:hi]
            snippet = " ".join(snippet.split())
            if lo > 0:
                snippet = "… " + snippet
            if hi < len(text):
                snippet = snippet + " …"
            if len(snippet) > 220:
                snippet = snippet[:217] + "…"
            hits.append({"label": label, "risk": risk, "snippet": snippet})
        return hits

    # -------------------------------------------------------------- doc check
    def load_file(self):
        path = filedialog.askopenfilename(
            title="Выберите документ",
            filetypes=[
                ("Все поддерживаемые", "*.txt *.docx *.pdf"),
                ("Текстовые файлы", "*.txt"),
                ("Word документы", "*.docx"),
                ("PDF документы", "*.pdf"),
                ("Все файлы", "*.*"),
            ],
        )
        if not path:
            return
        try:
            text = load_file_by_type(path)
            if not text or len(text) < 50:
                messagebox.showerror("Ошибка",
                    "Не удалось прочитать файл или он слишком короткий.")
                return

            self.current_file = path
            self.current_text = text
            st = self._doc_stats(text)
            self.file_label.config(
                text=(f"📄 {os.path.basename(path)}  |  "
                      f"{st['chars']:,} симв.  {st['words']:,} слов  "
                      f"{st['lines']} строк  {st['paras']} абзацев"),
                fg=self.BLUE,
            )
            self.check_btn.config(state=tk.NORMAL)
            self.save_btn.config(state=tk.DISABLED)
            self._show_doc_pane()
            self._display_load_preview(text, st)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при загрузке файла:\n{e}")

    def _display_load_preview(self, text: str, st: dict):
        """Показывает предварительный анализ при загрузке файла."""
        rt = self.result_text
        rt.delete(1.0, tk.END)
        self.content_title.config(text="📄 ДОКУМЕНТ ЗАГРУЖЕН — ПРЕДВАРИТЕЛЬНЫЙ АНАЛИЗ")

        # ── заголовок ─────────────────────────────────────────────────────
        rt.insert(tk.END, "=" * 80 + "\n", "header")
        rt.insert(tk.END, "  ДОКУМЕНТ ЗАГРУЖЕН\n", "success")
        rt.insert(tk.END, "=" * 80 + "\n\n", "header")

        # ── статистика ────────────────────────────────────────────────────
        rt.insert(tk.END, "СТАТИСТИКА ДОКУМЕНТА\n", "header")
        rt.insert(tk.END, "-" * 40 + "\n")
        rt.insert(tk.END, f"  Символов  : {st['chars']:>8,}\n")
        rt.insert(tk.END, f"  Слов      : {st['words']:>8,}\n")
        rt.insert(tk.END, f"  Строк     : {st['lines']:>8,}\n")
        rt.insert(tk.END, f"  Абзацев   : {st['paras']:>8,}\n")
        rt.insert(tk.END, f"  Предложений: {st['sents']:>7,}\n\n")

        # ── быстрый скан элементов ────────────────────────────────────────
        rt.insert(tk.END, "ЭКСПРЕСС-СКАНИРОВАНИЕ СТРУКТУРЫ\n", "header")
        rt.insert(tk.END, "-" * 40 + "\n")

        checks = self._quick_scan(text)
        found  = sum(1 for _, ok, _ in checks if ok)
        total  = len(checks)

        for name, ok, frag in checks:
            if ok:
                rt.insert(tk.END, "  [+] ", "success")
                rt.insert(tk.END, f"{name}\n", "normal")
                if frag and frag not in ("нет", "нет ссылок на приложения"):
                    rt.insert(tk.END, f"       → {frag}\n", "section")
            else:
                rt.insert(tk.END, "  [-] ", "error")
                rt.insert(tk.END, f"{name}\n", "normal")

        rt.insert(tk.END, "\n")
        pct = found * 100 // total
        bar = "█" * (found * 20 // total) + "░" * (20 - found * 20 // total)
        rt.insert(tk.END, f"  Соответствие: [{bar}] {pct}% ({found}/{total})\n\n",
                  "success" if pct >= 75 else "warning" if pct >= 50 else "error")

        # ── начало текста ─────────────────────────────────────────────────
        rt.insert(tk.END, "НАЧАЛО ДОКУМЕНТА (первые 600 символов)\n", "header")
        rt.insert(tk.END, "-" * 40 + "\n")
        preview = text[:600].strip()
        for line in preview.splitlines():
            stripped = line.strip()
            if not stripped:
                rt.insert(tk.END, "\n")
            elif stripped.isupper() and len(stripped) > 2:
                rt.insert(tk.END, f"  {stripped}\n", "section")
            else:
                rt.insert(tk.END, f"  {stripped}\n")
        if len(text) > 600:
            rt.insert(tk.END, "  ...\n")

        rt.insert(tk.END, "\n" + "=" * 80 + "\n")
        rt.insert(tk.END, "  Нажмите «Проверить структуру» для полного анализа\n", "section")
        rt.insert(tk.END, "=" * 80 + "\n")

    def check_document(self):
        if not self.current_text:
            messagebox.showwarning("Предупреждение", "Сначала загрузите документ")
            return

        self._show_doc_pane()
        self.check_btn.config(state=tk.DISABLED)
        self.load_btn.config(state=tk.DISABLED)
        self.progress.pack(side=tk.RIGHT, padx=10)
        self.progress.start(10)

        t = threading.Thread(target=self._check_thread, daemon=True)
        t.start()

    def _check_thread(self):
        try:
            val = self.validator.validate(self.current_text)
            mdl = None
            if self.model_available and self.trainer:
                mdl = self.trainer.predict(self.current_text)
            self.last_validation_result = val
            self.last_model_result = mdl
            self.root.after(0, self._display_results, val, mdl)
        except Exception as e:
            self.root.after(0, messagebox.showerror,
                            "Ошибка", f"Ошибка при проверке:\n{e}")
        finally:
            self.root.after(0, self._check_done)

    def _check_done(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.check_btn.config(state=tk.NORMAL)
        self.load_btn.config(state=tk.NORMAL)
        self.save_btn.config(state=tk.NORMAL)

    def _display_results(self, vr, mr):
        import re
        rt   = self.result_text
        text = self.current_text
        rt.delete(1.0, tk.END)
        self.content_title.config(text="📊 РЕЗУЛЬТАТЫ ПРОВЕРКИ СТРУКТУРЫ ПРИКАЗА")

        st     = self._doc_stats(text)
        checks = self._quick_scan(text)
        found  = sum(1 for _, ok, _ in checks if ok)
        total  = len(checks)
        pct    = found * 100 // total

        # ── шапка ─────────────────────────────────────────────────────────
        rt.insert(tk.END, "=" * 80 + "\n", "header")
        rt.insert(tk.END, "  РЕЗУЛЬТАТЫ ПРОВЕРКИ СТРУКТУРЫ ПРИКАЗА МЭР\n", "header")
        rt.insert(tk.END, "=" * 80 + "\n\n", "header")

        rt.insert(tk.END, f"  Файл    : {os.path.basename(self.current_file)}\n")
        rt.insert(tk.END, f"  Дата    : {datetime.now().strftime('%d.%m.%Y  %H:%M:%S')}\n")
        rt.insert(tk.END, f"  Размер  : {st['chars']:,} симв. / {st['words']:,} слов / "
                          f"{st['lines']} строк\n\n")

        # ── итог ──────────────────────────────────────────────────────────
        if vr["is_valid"] and pct >= 80:
            rt.insert(tk.END, "  [ИТОГ]  СТРУКТУРА ДОКУМЕНТА КОРРЕКТНА\n", "success")
        elif vr["total_errors"] == 0:
            rt.insert(tk.END, "  [ИТОГ]  Критических ошибок нет, есть замечания\n", "warning")
        else:
            rt.insert(tk.END, f"  [ИТОГ]  ОБНАРУЖЕНЫ НАРУШЕНИЯ СТРУКТУРЫ  "
                              f"(ошибок: {vr['total_errors']})\n", "error")

        bar = "█" * (pct // 5) + "░" * (20 - pct // 5)
        rt.insert(tk.END,
                  f"  Соответствие требованиям: [{bar}] {pct}%\n",
                  "success" if pct >= 75 else "warning" if pct >= 50 else "error")
        rt.insert(tk.END,
                  f"  Критических ошибок: {vr['total_errors']}   "
                  f"Предупреждений: {vr['total_warnings']}\n\n")

        # ── чек-лист по разделам ──────────────────────────────────────────
        rt.insert(tk.END, "=" * 80 + "\n", "header")
        rt.insert(tk.END, "  ЧЕК-ЛИСТ ОБЯЗАТЕЛЬНЫХ ЭЛЕМЕНТОВ\n", "header")
        rt.insert(tk.END, "=" * 80 + "\n")
        for name, ok, frag in checks:
            if ok:
                rt.insert(tk.END, "  [+] ", "success")
                rt.insert(tk.END, f"{name}\n")
                if frag and frag not in ("нет", "нет ссылок на приложения"):
                    rt.insert(tk.END, f"       «{frag}»\n", "section")
            else:
                rt.insert(tk.END, "  [-] ", "error")
                rt.insert(tk.END, f"{name}\n")
        rt.insert(tk.END, "\n")

        # ── разбор по разделам ────────────────────────────────────────────
        SECTION_ORDER = [
            "Шапка документа",
            "Заголовок",
            "Преамбула",
            "Распорядительная часть",
            "Приложения",
            "Прочие",
        ]
        SECTION_HELP = {
            "Шапка документа":
                "Должна содержать: полное наименование органа, вид документа «ПРИКАЗ»,\n"
                "    дату (ДД.ММ.ГГГГ), регистрационный номер (№ ...), место издания (г. ...).",
            "Заголовок":
                "Должен кратко отражать предмет приказа, начинаться с «О» или «Об»,\n"
                "    выноситься отдельной строкой перед преамбулой.",
            "Преамбула":
                "Указывает правовые основания издания приказа.\n"
                "    Начинается: «В соответствии с», «На основании», «Во исполнение» и т.п.\n"
                "    Должна содержать ссылку на НПА (с номером и датой).",
            "Распорядительная часть":
                "Начинается с отдельной строки «ПРИКАЗЫВАЮ:».\n"
                "    Содержит нумерованные пункты с конкретными распорядительными действиями.\n"
                "    Последний пункт — контроль за исполнением.",
            "Приложения":
                "Если упомянуты в тексте — должны быть оформлены отдельными блоками\n"
                "    с реквизитами «Приложение к приказу от ... № ...».",
            "Прочие":
                "Антикоррупционный анализ: поиск в тексте типовых коррупциогенных формулировок НПА\n"
                "    (усмотрение без критериев, размытые сроки, широкие отсылки и т.п.).\n"
                "    Полный методический перечень — в сохранённом отчёте (конец файла).",
        }

        # группируем ошибки/предупреждения по разделу
        from collections import defaultdict
        err_by_sec  = defaultdict(list)
        warn_by_sec = defaultdict(list)
        for e in vr["errors"]:
            err_by_sec[e.section].append(e)
        for w in vr["warnings"]:
            warn_by_sec[w.section].append(w)

        all_sections = set(err_by_sec) | set(warn_by_sec)
        ordered = [s for s in SECTION_ORDER if s in all_sections or s in SECTION_HELP]

        corr_hits = self._scan_corruption_risks(text)

        rt.insert(tk.END, "=" * 80 + "\n", "header")
        rt.insert(tk.END, "  ДЕТАЛЬНЫЙ АНАЛИЗ ПО РАЗДЕЛАМ\n", "header")
        rt.insert(tk.END, "=" * 80 + "\n\n")

        for sec in ordered:
            if sec == "Прочие":
                if corr_hits:
                    marker, tag = "[ЗАМЕЧАНИЯ]", "warning"
                else:
                    marker, tag = "[ OK ]     ", "success"
                rt.insert(tk.END, f"  {marker}  {sec}\n", tag)
                rt.insert(tk.END, f"    Требования: {SECTION_HELP[sec]}\n", "section")
                for hi in corr_hits:
                    rt.insert(tk.END, f"    ~ Возможная коррупциогенная конструкция\n", "warning")
                    rt.insert(tk.END, f"      «{hi['label']}» — риск: {hi['risk']}\n", "corr_note")
                    rt.insert(tk.END, f"      Фрагмент: {hi['snippet']}\n", "normal")
                if not corr_hits:
                    rt.insert(tk.END, "    Типовые формулировки из перечня в тексте не обнаружены (автопоиск).\n",
                              "success")
                rt.insert(tk.END, "\n")
                continue

            errs  = err_by_sec.get(sec, [])
            warns = warn_by_sec.get(sec, [])
            if errs:
                marker, tag = "[ ОШИБКИ ]", "error"
            elif warns:
                marker, tag = "[ЗАМЕЧАНИЯ]", "warning"
            else:
                marker, tag = "[ OK ]     ", "success"

            rt.insert(tk.END, f"  {marker}  {sec}\n", tag)
            if sec in SECTION_HELP:
                rt.insert(tk.END, f"    Требования: {SECTION_HELP[sec]}\n", "section")

            for e in errs:
                rt.insert(tk.END, f"    ! {e.error_type}\n", "error")
                rt.insert(tk.END, f"      {e.description}\n")
            for w in warns:
                rt.insert(tk.END, f"    ~ {w.error_type}\n", "warning")
                rt.insert(tk.END, f"      {w.description}\n")
            rt.insert(tk.END, "\n")

        # ── извлечённые фрагменты из текста ──────────────────────────────
        rt.insert(tk.END, "=" * 80 + "\n", "header")
        rt.insert(tk.END, "  ИЗВЛЕЧЁННЫЕ РЕКВИЗИТЫ ДОКУМЕНТА\n", "header")
        rt.insert(tk.END, "=" * 80 + "\n")

        head20 = "\n".join(text.splitlines()[:25])
        fragments = {
            "Наименование органа": re.search(
                r'(МИНИСТЕРСТВО[^\n]+|Минэкономразвития[^\n]+)', head20, re.I),
            "Дата приказа": re.search(
                r'(?:от\s+)?(\d{2}\.\d{2}\.\d{4}|«?\d{1,2}»?\s+[а-яё]+\s+\d{4})', text, re.I),
            "Номер приказа": re.search(r'(№\s*[\d\w/-]+)', text),
            "Место издания": re.search(r'(г\.\s*[А-ЯЁ][а-яё-]+)', head20),
            "Заголовок «О ...»": re.search(
                r'^\s*(О[бб]?\s+[А-ЯЁа-яё].{10,80})', text, re.MULTILINE),
            "Распорядительное слово": re.search(r'(ПРИКАЗЫВАЮ\s*:?)', text, re.I),
            "Подписант": re.search(
                r'([Мм]инистр\s+[^\n]{5,40}|[Зз]ам\w*\s+министра[^\n]{0,30})', text),
        }
        for label, m in fragments.items():
            val = m.group(1).strip()[:70] if m else "—  не найдено"
            tag = "section" if m else "error"
            rt.insert(tk.END, f"  {label:<28}: ", "normal")
            rt.insert(tk.END, f"{val}\n", tag)
        rt.insert(tk.END, "\n")

        # ── нейронная модель ──────────────────────────────────────────────
        if mr:
            rt.insert(tk.END, "=" * 80 + "\n", "header")
            rt.insert(tk.END, "  ОЦЕНКА НЕЙРОННОЙ МОДЕЛЬЮ (rubert-tiny2)\n", "header")
            rt.insert(tk.END, "=" * 80 + "\n")
            conf = mr["confidence"] * 100
            if mr["has_errors"]:
                rt.insert(tk.END, f"  Вердикт  : ОШИБКИ ОБНАРУЖЕНЫ  (уверенность {conf:.1f}%)\n",
                          "error")
            else:
                rt.insert(tk.END, f"  Вердикт  : НАРУШЕНИЙ НЕ ВЫЯВЛЕНО  (уверенность {conf:.1f}%)\n",
                          "success")
            no_p  = mr["probabilities"]["no_errors"]  * 100
            has_p = mr["probabilities"]["has_errors"] * 100
            rt.insert(tk.END, f"  P(корректен)   : {no_p:.1f}%\n")
            rt.insert(tk.END, f"  P(есть ошибки) : {has_p:.1f}%\n")
            if mr["pattern_errors"]:
                rt.insert(tk.END, "\n  Ошибки по шаблонам:\n", "section")
                for cat, errs in mr["pattern_errors"].items():
                    rt.insert(tk.END, f"    - {cat}: {len(errs)} случаев\n")
            rt.insert(tk.END, "\n")

        # ── рекомендации ──────────────────────────────────────────────────
        rec = []
        for e in vr["errors"]:
            rec.append(("ОБЯЗАТЕЛЬНО", e.section, e.error_type, e.description))
        for w in vr["warnings"]:
            rec.append(("ЖЕЛАТЕЛЬНО", w.section, w.error_type, w.description))
        if mr and mr["has_errors"]:
            rec.append(("ОБЯЗАТЕЛЬНО", "Модель", "Документ признан ошибочным",
                        "Устраните нарушения и повторите проверку."))

        if rec:
            rt.insert(tk.END, "=" * 80 + "\n", "header")
            rt.insert(tk.END, "  РЕКОМЕНДАЦИИ ПО УСТРАНЕНИЮ НАРУШЕНИЙ\n", "header")
            rt.insert(tk.END, "=" * 80 + "\n")
            oblig = [(s, t, d) for p, s, t, d in rec if p == "ОБЯЗАТЕЛЬНО"]
            optl  = [(s, t, d) for p, s, t, d in rec if p == "ЖЕЛАТЕЛЬНО"]
            if oblig:
                rt.insert(tk.END, "  Критические (устранить обязательно):\n", "error")
                for i, (s, t, d) in enumerate(oblig, 1):
                    rt.insert(tk.END, f"  {i}. [{s}] {t}\n", "error")
                    rt.insert(tk.END, f"     {d}\n")
                rt.insert(tk.END, "\n")
            if optl:
                rt.insert(tk.END, "  Рекомендуемые улучшения:\n", "warning")
                for i, (s, t, d) in enumerate(optl, 1):
                    rt.insert(tk.END, f"  {i}. [{s}] {t}\n", "warning")
                    rt.insert(tk.END, f"     {d}\n")
            rt.insert(tk.END, "\n")

        # ── финал ─────────────────────────────────────────────────────────
        rt.insert(tk.END, "=" * 80 + "\n", "header")
        conclusion = (
            "ЗАКЛЮЧЕНИЕ: документ соответствует требованиям оформления приказов МЭР."
            if vr["is_valid"] and pct >= 75 else
            "ЗАКЛЮЧЕНИЕ: документ требует доработки перед направлением на согласование."
        )
        rt.insert(tk.END, f"  {conclusion}\n",
                  "success" if vr["is_valid"] and pct >= 75 else "error")
        rt.insert(tk.END, "=" * 80 + "\n")

    # ----------------------------------------------------------------- report
    def save_report(self):
        if not self.last_validation_result:
            messagebox.showwarning("Предупреждение",
                                   "Сначала выполните проверку документа")
            return

        path = filedialog.asksaveasfilename(
            title="Сохранить отчёт",
            defaultextension=".txt",
            filetypes=[("Текстовый файл", "*.txt"), ("Все файлы", "*.*")],
            initialfile=f"отчет_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8-sig", newline="\n") as f:
                f.write(self._generate_report())
            messagebox.showinfo("Успешно",
                f"Отчёт сохранён:\n{os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при сохранении:\n{e}")

    def _generate_report(self):
        import re
        vr   = self.last_validation_result
        mr   = self.last_model_result
        text = self.current_text
        st   = self._doc_stats(text)
        checks = self._quick_scan(text)
        found  = sum(1 for _, ok, _ in checks if ok)
        total  = len(checks)
        pct    = found * 100 // total
        now    = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
        bar    = "#" * (pct // 5) + "." * (20 - pct // 5)

        L = []
        def h(s=""):  L.append(s)
        def hr(): L.append("=" * 80)
        def sr(): L.append("-" * 60)

        hr(); h("  ОФИЦИАЛЬНЫЙ ОТЧЁТ О ПРОВЕРКЕ СТРУКТУРЫ ПРИКАЗА МЭР")
        h("  Министерство экономического развития Российской Федерации")
        hr(); h()
        h(f"  Дата проверки  : {now}")
        h(f"  Файл           : {os.path.basename(self.current_file)}")
        h(f"  Символов       : {st['chars']:,}   Слов: {st['words']:,}   "
          f"Строк: {st['lines']}   Абзацев: {st['paras']}")
        h()

        # итог
        hr()
        h("  СВОДНОЕ ЗАКЛЮЧЕНИЕ")
        hr()
        verdict = ("СТРУКТУРА СООТВЕТСТВУЕТ ТРЕБОВАНИЯМ" if vr["is_valid"] and pct >= 75
                   else "ОБНАРУЖЕНЫ НАРУШЕНИЯ СТРУКТУРЫ ДОКУМЕНТА")
        h(f"  Вердикт        : {verdict}")
        h(f"  Соответствие   : [{bar}] {pct}% ({found}/{total} элементов в норме)")
        h(f"  Критических ошибок : {vr['total_errors']}")
        h(f"  Предупреждений     : {vr['total_warnings']}")
        h()
        h("  + В отчёт включён раздел «Прочие»: краткий итог — в «Детальном анализе по разделам»;")
        h("    полный методический перечень и повтор результатов поиска — перед итоговым заключением.")
        h()

        # чек-лист
        hr(); h("  ЧЕК-ЛИСТ ОБЯЗАТЕЛЬНЫХ ЭЛЕМЕНТОВ ПРИКАЗА"); hr()
        for name, ok, frag in checks:
            mark = "[+]" if ok else "[-]"
            h(f"  {mark}  {name}")
            if ok and frag and frag not in ("нет", "нет ссылок на приложения"):
                h(f"        > {frag}")
        h()

        # реквизиты
        hr(); h("  ИЗВЛЕЧЁННЫЕ РЕКВИЗИТЫ ДОКУМЕНТА"); hr()
        head20 = "\n".join(text.splitlines()[:25])
        fragments = {
            "Наименование органа": re.search(
                r'(МИНИСТЕРСТВО[^\n]+|Минэкономразвития[^\n]+)', head20, re.I),
            "Дата приказа": re.search(
                r'(?:от\s+)?(\d{2}\.\d{2}\.\d{4}|«?\d{1,2}»?\s+[а-яё]+\s+\d{4})', text, re.I),
            "Номер приказа": re.search(r'(№\s*[\d\w/-]+)', text),
            "Место издания": re.search(r'(г\.\s*[А-ЯЁ][а-яё-]+)', head20),
            "Заголовок":     re.search(r'^\s*(О[бб]?\s+[А-ЯЁа-яё].{10,80})',
                                        text, re.MULTILINE),
            "Распорядительное слово": re.search(r'(ПРИКАЗЫВАЮ\s*:?)', text, re.I),
            "Подписант": re.search(
                r'([Мм]инистр\s+[^\n]{5,40}|[Зз]ам\w*\s+министра[^\n]{0,30})', text),
        }
        for label, m in fragments.items():
            val = m.group(1).strip()[:70] if m else "— не найдено"
            h(f"  {label:<26}: {val}")
        h()

        # детали по разделам
        hr(); h("  ДЕТАЛЬНЫЙ АНАЛИЗ ПО РАЗДЕЛАМ"); hr()
        h("  Порядок: шапка документа, заголовок, преамбула, распорядительная часть,")
        h("           приложения, затем раздел ПРОЧИЕ (антикоррупционный анализ текста).")
        h()
        SECTION_ORDER = ["Шапка документа", "Заголовок", "Преамбула",
                         "Распорядительная часть", "Приложения", "Прочие"]
        from collections import defaultdict
        eb = defaultdict(list); wb = defaultdict(list)
        for e in vr["errors"]:   eb[e.section].append(e)
        for w in vr["warnings"]: wb[w.section].append(w)

        HELP = {
            "Шапка документа":
                "Наименование органа, ПРИКАЗ, дата, номер, место издания.",
            "Заголовок":
                "Начинается с «О» / «Об», краткий предмет приказа.",
            "Преамбула":
                "«В соответствии с ...», «На основании ...» + ссылка на НПА.",
            "Распорядительная часть":
                "«ПРИКАЗЫВАЮ:» + нумерованные пункты + пункт о контроле.",
            "Приложения":
                "Реквизиты «Приложение к приказу от ... № ...».",
            "Прочие":
                "Антикоррупционный анализ типовых формулировок НПА (автопоиск по тексту).",
        }
        corr_hits_rep = self._scan_corruption_risks(text)
        for sec in SECTION_ORDER:
            if sec == "Прочие":
                status = "ЗАМЕЧАНИЯ" if corr_hits_rep else "OK"
                h(f"  [{status:<9}]  {sec}")
                h(f"    Требования: {HELP[sec]}")
                for hi in corr_hits_rep:
                    h("    ~ Возможная коррупциогенная конструкция")
                    h(f"      «{hi['label']}» — риск: {hi['risk']}")
                    h(f"      Фрагмент: {hi['snippet']}")
                if not corr_hits_rep:
                    h("    Типовые формулировки из перечня в тексте не обнаружены (автопоиск).")
                h()
                continue
            errs = eb.get(sec, []); warns = wb.get(sec, [])
            status = "ОШИБКИ" if errs else ("ЗАМЕЧАНИЯ" if warns else "OK")
            h(f"  [{status:<9}]  {sec}")
            if sec in HELP: h(f"    Требования: {HELP[sec]}")
            for e in errs:   h(f"    ! {e.error_type}\n      {e.description}")
            for w in warns:  h(f"    ~ {w.error_type}\n      {w.description}")
            if not errs and not warns: h("    Нарушений не выявлено.")
            h()

        # нейронная модель
        if mr:
            hr(); h("  ОЦЕНКА НЕЙРОННОЙ МОДЕЛЬЮ (rubert-tiny2)"); hr()
            conf = mr["confidence"] * 100
            v2 = "ОШИБКИ ОБНАРУЖЕНЫ" if mr["has_errors"] else "НАРУШЕНИЙ НЕ ВЫЯВЛЕНО"
            h(f"  Вердикт модели  : {v2}  (уверенность {conf:.1f}%)")
            h(f"  P(корректен)    : {mr['probabilities']['no_errors']*100:.1f}%")
            h(f"  P(есть ошибки)  : {mr['probabilities']['has_errors']*100:.1f}%")
            if mr["pattern_errors"]:
                h("  Ошибки по шаблонам:")
                for cat, errs in mr["pattern_errors"].items():
                    h(f"    - {cat}: {len(errs)} случаев")
            h()

        # рекомендации
        all_rec = ([(True,  e.section, e.error_type, e.description) for e in vr["errors"]] +
                   [(False, w.section, w.error_type, w.description) for w in vr["warnings"]])
        if all_rec:
            hr(); h("  РЕКОМЕНДАЦИИ ПО УСТРАНЕНИЮ НАРУШЕНИЙ"); hr()
            crit = [(s,t,d) for critical,s,t,d in all_rec if critical]
            opts = [(s,t,d) for critical,s,t,d in all_rec if not critical]
            if crit:
                h("  Обязательные меры:")
                for i,(s,t,d) in enumerate(crit,1):
                    h(f"  {i}. [{s}] {t}"); h(f"     {d}")
                h()
            if opts:
                h("  Рекомендуемые улучшения:")
                for i,(s,t,d) in enumerate(opts,1):
                    h(f"  {i}. [{s}] {t}"); h(f"     {d}")
            h()

        try:
            for line in self._corruption_report_lines(text):
                h(line)
        except Exception as ex:
            hr()
            h("  Ошибка при добавлении полного приложения «Прочие» в отчёт:")
            h(f"    {ex}")
            h("  Краткий раздел «Прочие» должен быть выше — в «Детальном анализе по разделам».")
            hr()

        hr()
        h("  ИТОГОВОЕ ЗАКЛЮЧЕНИЕ")
        h("  " + ("Документ соответствует требованиям оформления приказов МЭР."
                  if vr["is_valid"] and pct >= 75 else
                  "Документ требует доработки перед направлением на согласование."))
        hr()
        h("  Подготовлено автоматической системой проверки приказов МЭР")
        h(f"  Дата: {now}")
        hr()
        out = "\n".join(L)
        if "Прочие" not in out:
            out += (
                "\n\n"
                + "=" * 80
                + "\n  ВНИМАНИЕ: в отчёте отсутствует раздел «Прочие» (ошибка сборки). "
                "Обновите программу до последней версии gui_app.py.\n"
                + "=" * 80
            )
        return out

    def _corruption_report_lines(self, text: str):
        """Строки раздела «Прочие» для текстового отчёта (единая сборка)."""
        lines = []
        sep = "=" * 80
        dash = "-" * 60

        lines.append("")
        lines.append(sep)
        lines.append("  РАЗДЕЛ «ПРОЧИЕ» — АНТИКОРРУПЦИОННЫЙ АНАЛИЗ ФОРМУЛИРОВОК НПА")
        lines.append(sep)
        lines.append("")
        lines.append("  Ниже — методический ориентир и результат автоматического поиска в тексте документа.")
        lines.append("  Поиск не заменяет экспертизу; возможны ложные срабатывания.")
        lines.append("")
        for ref_line in CORRUPTION_REFERENCE_TEXT.splitlines():
            lines.append("  " + ref_line if ref_line.strip() else "")
        lines.append("")
        lines.append(dash)
        lines.append("  Результат поиска в тексте (категория «Прочие»)")
        lines.append(dash)
        corr_hits = self._scan_corruption_risks(text)
        if corr_hits:
            lines.append(f"  Обнаружено возможных совпадений: {len(corr_hits)}.")
            lines.append("  Рекомендуется ручная проверка контекста.")
            lines.append("")
            for i, item in enumerate(corr_hits, 1):
                lines.append(f"  {i}. «{item['label']}»")
                lines.append(f"     Риск: {item['risk']}")
                lines.append(f"     Фрагмент: {item['snippet']}")
                lines.append("")
        else:
            lines.append("  Типовые конструкции из перечня в тексте не обнаружены (по правилам автоматического поиска).")
            lines.append("  Иные размытые формулировки возможны — при необходимости проведите антикоррупционную экспертизу.")
            lines.append("")
        return lines

    # ------------------------------------------------------------------ clear
    def clear_results(self):
        self.result_text.delete(1.0, tk.END)
        self.file_label.config(text="📄 Файл не загружен", fg=self.BLUE)
        self.current_file = None
        self.current_text = None
        self.last_validation_result = None
        self.last_model_result = None
        self.check_btn.config(state=tk.DISABLED)
        self.save_btn.config(state=tk.DISABLED)
        self._show_doc_pane()


def main():
    root = tk.Tk()
    app = OrderCheckerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
