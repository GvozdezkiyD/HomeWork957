# -*- coding: utf-8 -*-
"""
Модуль для проверки структуры приказа
Проверяет наличие обязательных реквизитов и правильность оформления
"""

import re
from typing import Dict, List, Tuple
from dataclasses import dataclass


@dataclass
class ValidationError:
    """Класс для представления ошибки валидации"""
    section: str  # Раздел документа
    error_type: str  # Тип ошибки
    description: str  # Описание ошибки
    line_number: int = None  # Номер строки (если есть)


class OrderStructureValidator:
    """
    Класс для проверки структуры приказа на соответствие требованиям
    """
    
    def __init__(self):
        """Инициализация валидатора"""
        self.errors: List[ValidationError] = []
        self.warnings: List[ValidationError] = []
    
    def validate(self, text: str) -> Dict:
        """
        Основной метод валидации структуры приказа
        
        Args:
            text: текст приказа
            
        Returns:
            словарь с результатами проверки
        """
        self.errors = []
        self.warnings = []
        
        # Разбиваем текст на строки для анализа
        lines = text.split('\n')
        
        # 1. Проверка реквизитов "шапки"
        self._check_header(text, lines)
        
        # 2. Проверка заголовка к тексту
        self._check_title(text, lines)
        
        # 3. Проверка преамбулы
        self._check_preamble(text, lines)
        
        # 4. Проверка распорядительной части
        self._check_directive_part(text, lines)
        
        # 5. Проверка приложений
        self._check_attachments(text, lines)
        
        return {
            'is_valid': len(self.errors) == 0,
            'errors': self.errors,
            'warnings': self.warnings,
            'total_errors': len(self.errors),
            'total_warnings': len(self.warnings)
        }
    
    def _check_header(self, text: str, lines: List[str]):
        """
        Проверка реквизитов "шапки" документа
        
        Проверяет:
        - Наименование органа (министерства)
        - Тип документа (ПРИКАЗ)
        - Дата (формат ДД.ММ.ГГГГ)
        - Номер (№ и значение)
        - Место издания (г. <город>)
        """
        # Объединяем первые 20 строк для анализа шапки
        header_text = '\n'.join(lines[:20])
        
        # 1. Проверка наименования органа
        ministry_patterns = [
            r'министерство',
            r'минэкономразвития',
            r'федеральн',
            r'департамент',
            r'управление'
        ]
        
        has_ministry = any(re.search(pattern, header_text, re.IGNORECASE) for pattern in ministry_patterns)
        
        if not has_ministry:
            self.errors.append(ValidationError(
                section="Шапка документа",
                error_type="Отсутствует наименование органа",
                description="В начале документа не найдено наименование организации (министерства, департамента и т.п.)"
            ))
        
        # 2. Проверка типа документа "ПРИКАЗ"
        has_order_word = re.search(r'^ПРИКАЗ$', header_text, re.MULTILINE | re.IGNORECASE)
        if not has_order_word:
            has_order_word = re.search(r'ПРИКАЗ', header_text, re.IGNORECASE)
            if has_order_word:
                self.warnings.append(ValidationError(
                    section="Шапка документа",
                    error_type="Некорректное оформление слова ПРИКАЗ",
                    description="Слово 'ПРИКАЗ' должно быть отдельной строкой в верхнем регистре"
                ))
            else:
                self.errors.append(ValidationError(
                    section="Шапка документа",
                    error_type="Отсутствует тип документа",
                    description="Не найдено слово 'ПРИКАЗ' в начале документа"
                ))
        
        # 3. Проверка даты
        date_patterns = [
            r'\d{2}\.\d{2}\.\d{4}',  # ДД.ММ.ГГГГ
            r'\d{2}\s+[а-яё]+\s+\d{4}',  # ДД месяц ГГГГ
            r'«\d{1,2}»\s+[а-яё]+\s+\d{4}'  # «ДД» месяц ГГГГ
        ]
        
        has_date = any(re.search(pattern, header_text, re.IGNORECASE) for pattern in date_patterns)
        
        if not has_date:
            self.errors.append(ValidationError(
                section="Шапка документа",
                error_type="Отсутствует дата",
                description="Не найдена дата издания приказа в формате ДД.ММ.ГГГГ или словесном формате"
            ))
        
        # 4. Проверка номера
        number_pattern = r'№\s*\d+'
        has_number = re.search(number_pattern, header_text)
        
        if not has_number:
            self.errors.append(ValidationError(
                section="Шапка документа",
                error_type="Отсутствует номер",
                description="Не найден номер приказа (формат: № <цифры>)"
            ))
        
        # 5. Проверка места издания (необязательно для зарегистрированных)
        city_pattern = r'г\.\s*[А-ЯЁ][а-яё-]+'
        has_city = re.search(city_pattern, header_text)
        
        # Если документ зарегистрирован в Минюсте, место издания может отсутствовать
        is_registered = re.search(r'Зарегистрировано в Минюсте', header_text, re.IGNORECASE)
        
        if not has_city and not is_registered:
            self.warnings.append(ValidationError(
                section="Шапка документа",
                error_type="Отсутствует место издания",
                description="Не найдено место издания (формат: г. <Город>). Рекомендуется указывать для незарегистрированных документов."
            ))
    
    def _check_title(self, text: str, lines: List[str]):
        """
        Проверка заголовка к тексту
        
        Заголовок должен:
        - Начинаться с "О" или "Об"
        - Содержать минимум 2-3 слова
        """
        # Ищем заголовок после шапки (обычно после слова ПРИКАЗ и до преамбулы)
        title_patterns = [
            r'(?:^|\n)\s*О\s+[а-яёА-ЯЁ\s]{10,}',  # О ...
            r'(?:^|\n)\s*Об\s+[а-яёА-ЯЁ\s]{10,}'  # Об ...
        ]
        
        has_title = False
        title_match = None
        
        for pattern in title_patterns:
            title_match = re.search(pattern, text, re.IGNORECASE)
            if title_match:
                has_title = True
                break
        
        if not has_title:
            self.warnings.append(ValidationError(
                section="Заголовок",
                error_type="Отсутствует заголовок",
                description="Не найден заголовок к тексту, начинающийся с 'О' или 'Об'. Рекомендуется для структурированных приказов."
            ))
        else:
            # Проверяем длину заголовка
            title_text = title_match.group(0) if title_match else ""
            words = title_text.split()
            
            if len(words) < 3:
                self.errors.append(ValidationError(
                    section="Заголовок",
                    error_type="Заголовок слишком короткий",
                    description=f"Заголовок содержит менее 3 слов: '{title_text.strip()}'"
                ))
    
    def _check_preamble(self, text: str, lines: List[str]):
        """
        Проверка преамбулы (основание издания)
        
        Преамбула должна содержать:
        - Ключевые слова: "В соответствии с", "На основании", "Во исполнение", "Руководствуясь"
        - Ссылку на НПА (номер, дата, название)
        """
        preamble_keywords = [
            r'В\s+соответствии\s+с',
            r'На\s+основании',
            r'Во\s+исполнение',
            r'Руководствуясь',
            r'В\s+целях',
            r'В\s+связи\s+с'
        ]
        
        has_preamble = any(re.search(pattern, text, re.IGNORECASE) for pattern in preamble_keywords)
        
        if not has_preamble:
            self.errors.append(ValidationError(
                section="Преамбула",
                error_type="Отсутствует преамбула",
                description="Не найдено основание издания приказа (должны быть формулы: 'В соответствии с', 'На основании', 'Руководствуясь' и т.п.)"
            ))
        else:
            # Проверяем наличие ссылки на НПА
            npa_patterns = [
                r'[Фф]едеральн\w+\s+закон\w*',
                r'[Пп]остановлени\w+\s+[Пп]равительства',
                r'[Пп]риказ\w*\s+.*№\s*\d+',
                r'№\s*\d+[-\w]*\s+от\s+\d{2}\.\d{2}\.\d{4}',
                r'от\s+\d{2}\.\d{2}\.\d{4}\s+№\s*\d+'
            ]
            
            has_npa_reference = any(re.search(pattern, text, re.IGNORECASE) for pattern in npa_patterns)
            
            if not has_npa_reference:
                self.warnings.append(ValidationError(
                    section="Преамбула",
                    error_type="Отсутствует ссылка на НПА",
                    description="В преамбуле желательно указать ссылку на нормативно-правовой акт (закон, постановление, приказ с номером и датой)"
                ))
    
    def _check_directive_part(self, text: str, lines: List[str]):
        """
        Проверка распорядительной части
        
        Должна содержать:
        - Слово "ПРИКАЗЫВАЮ:" отдельной строкой
        - Нумерованные пункты (1., 2., 3., ...)
        - Действия (глаголы: утвердить, установить, назначить и т.п.)
        - Пункт про контроль исполнения
        """
        # 1. Проверка слова "ПРИКАЗЫВАЮ"
        prikazyvayu_pattern = r'(?:^|\n)\s*ПРИКАЗЫВАЮ\s*:'
        has_prikazyvayu = re.search(prikazyvayu_pattern, text, re.IGNORECASE)
        
        if not has_prikazyvayu:
            self.errors.append(ValidationError(
                section="Распорядительная часть",
                error_type="Отсутствует слово ПРИКАЗЫВАЮ",
                description="Распорядительная часть должна начинаться со слова 'ПРИКАЗЫВАЮ:' отдельной строкой"
            ))
            return  # Если нет ПРИКАЗЫВАЮ, дальнейшие проверки неактуальны
        
        # Извлекаем текст после ПРИКАЗЫВАЮ
        prikazyvayu_pos = has_prikazyvayu.end()
        directive_text = text[prikazyvayu_pos:]
        
        # 2. Проверка нумерованных пунктов
        # Ищем пункты вида "1.", "2.", "3." и т.д.
        numbered_items = re.findall(r'(?:^|\n)\s*(\d+)\.\s+([^\n]+)', directive_text, re.MULTILINE)
        
        if not numbered_items:
            self.errors.append(ValidationError(
                section="Распорядительная часть",
                error_type="Отсутствуют нумерованные пункты",
                description="После 'ПРИКАЗЫВАЮ:' должны следовать нумерованные пункты (1., 2., 3., ...)"
            ))
            return
        
        # 3. Проверка последовательности нумерации
        numbers = [int(num) for num, _ in numbered_items]
        expected_numbers = list(range(1, len(numbers) + 1))
        
        if numbers != expected_numbers:
            missing = set(expected_numbers) - set(numbers)
            duplicates = [num for num in numbers if numbers.count(num) > 1]
            
            error_desc = "Нарушена последовательность нумерации пунктов. "
            if missing:
                error_desc += f"Пропущены номера: {sorted(missing)}. "
            if duplicates:
                error_desc += f"Повторяющиеся номера: {sorted(set(duplicates))}."
            
            self.errors.append(ValidationError(
                section="Распорядительная часть",
                error_type="Некорректная нумерация",
                description=error_desc
            ))
        
        # 4. Проверка наличия действий (глаголов)
        action_verbs = [
            r'утвердить',
            r'установить',
            r'назначить',
            r'признать',
            r'внести',
            r'обеспечить',
            r'создать',
            r'организовать',
            r'возложить',
            r'определить',
            r'провести'
        ]
        
        empty_items = []
        for num, content in numbered_items:
            content_lower = content.lower()
            has_action = any(re.search(verb, content_lower) for verb in action_verbs)
            
            if len(content.strip()) < 10:  # Слишком короткий пункт
                empty_items.append(num)
            elif not has_action:
                self.warnings.append(ValidationError(
                    section="Распорядительная часть",
                    error_type=f"Пункт {num} не содержит действия",
                    description=f"Пункт {num} возможно не содержит распорядительного глагола (утвердить, установить, назначить и т.п.): '{content[:50]}...'"
                ))
        
        if empty_items:
            self.errors.append(ValidationError(
                section="Распорядительная часть",
                error_type="Пустые пункты",
                description=f"Следующие пункты слишком короткие или пустые: {empty_items}"
            ))
        
        # 5. Проверка пункта о контроле исполнения
        control_patterns = [
            r'контроль\s+за\s+исполнением',
            r'контроль\s+исполнения',
            r'возложить\s+контроль'
        ]
        
        has_control = any(re.search(pattern, directive_text, re.IGNORECASE) for pattern in control_patterns)
        
        if not has_control:
            self.warnings.append(ValidationError(
                section="Распорядительная часть",
                error_type="Отсутствует пункт о контроле",
                description="Рекомендуется включить пункт о контроле за исполнением приказа (например: 'Контроль за исполнением настоящего приказа возложить на...')"
            ))
    
    def _check_attachments(self, text: str, lines: List[str]):
        """
        Проверка приложений
        
        Если в тексте упоминаются приложения:
        - Должен быть раздел с приложением
        - Приложение должно иметь реквизиты
        - Номера приложений должны быть последовательными
        """
        # Проверяем упоминания приложений в основном тексте
        attachment_mentions = re.findall(
            r'[Пп]риложени[еия]\s*(?:№\s*)?(\d+)?',
            text,
            re.IGNORECASE
        )
        
        if not attachment_mentions:
            # Упоминаний нет - проверка не требуется
            return
        
        # Проверяем наличие блоков приложений
        attachment_blocks = re.findall(
            r'(?:^|\n)\s*Приложение\s*(?:№\s*)?(\d+)?\s*\n.*?к\s+приказу',
            text,
            re.IGNORECASE | re.DOTALL
        )
        
        if len(attachment_mentions) > 0 and len(attachment_blocks) == 0:
            self.errors.append(ValidationError(
                section="Приложения",
                error_type="Отсутствуют приложения",
                description=f"В тексте упоминаются приложения ({len(attachment_mentions)} раз), но разделы приложений не найдены"
            ))
        elif len(attachment_blocks) > 0:
            # Проверяем реквизиты приложений
            for i, block in enumerate(attachment_blocks, 1):
                # Проверяем наличие номера, даты в реквизитах
                has_requisites = re.search(
                    r'к\s+приказу.*?(?:от|№)',
                    text[text.find(f'Приложение'):text.find(f'Приложение') + 200],
                    re.IGNORECASE | re.DOTALL
                )
                
                if not has_requisites:
                    self.warnings.append(ValidationError(
                        section="Приложения",
                        error_type=f"Неполные реквизиты приложения {i}",
                        description=f"Приложение {i} должно содержать реквизиты: 'Приложение к приказу от <дата> № <номер>'"
                    ))
            
            # Проверяем последовательность номеров приложений (если они есть)
            attachment_numbers = [num for num in attachment_mentions if num]
            if attachment_numbers:
                attachment_numbers = [int(n) for n in attachment_numbers if n.isdigit()]
                if attachment_numbers:
                    expected = list(range(1, max(attachment_numbers) + 1))
                    if sorted(set(attachment_numbers)) != expected:
                        self.warnings.append(ValidationError(
                            section="Приложения",
                            error_type="Некорректная нумерация приложений",
                            description=f"Номера приложений должны быть последовательными: найдены {sorted(set(attachment_numbers))}"
                        ))
    
    def get_report(self, validation_result: Dict) -> str:
        """
        Формирует текстовый отчет о результатах проверки
        
        Args:
            validation_result: результат проверки
            
        Returns:
            форматированный отчет
        """
        report = []
        report.append("=" * 80)
        report.append("ОТЧЕТ О ПРОВЕРКЕ СТРУКТУРЫ ПРИКАЗА")
        report.append("=" * 80)
        
        if validation_result['is_valid']:
            report.append("\n✓ СТРУКТУРА ДОКУМЕНТА КОРРЕКТНА")
            report.append("\nДокумент соответствует всем обязательным требованиям.")
        else:
            report.append("\n✗ ОБНАРУЖЕНЫ ОШИБКИ В СТРУКТУРЕ ДОКУМЕНТА")
            report.append(f"\nВсего ошибок: {validation_result['total_errors']}")
            report.append(f"Всего предупреждений: {validation_result['total_warnings']}")
        
        # Ошибки
        if validation_result['errors']:
            report.append("\n" + "=" * 80)
            report.append("КРИТИЧЕСКИЕ ОШИБКИ:")
            report.append("=" * 80)
            
            for i, error in enumerate(validation_result['errors'], 1):
                report.append(f"\n{i}. [{error.section}] {error.error_type}")
                report.append(f"   {error.description}")
        
        # Предупреждения
        if validation_result['warnings']:
            report.append("\n" + "=" * 80)
            report.append("ПРЕДУПРЕЖДЕНИЯ:")
            report.append("=" * 80)
            
            for i, warning in enumerate(validation_result['warnings'], 1):
                report.append(f"\n{i}. [{warning.section}] {warning.error_type}")
                report.append(f"   {warning.description}")
        
        report.append("\n" + "=" * 80)
        
        return '\n'.join(report)


def validate_order(text: str) -> Dict:
    """
    Вспомогательная функция для проверки одного приказа
    
    Args:
        text: текст приказа
        
    Returns:
        результат проверки
    """
    validator = OrderStructureValidator()
    return validator.validate(text)


if __name__ == "__main__":
    # Тестирование модуля
    test_text = """
МИНИСТЕРСТВО ЭКОНОМИЧЕСКОГО РАЗВИТИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ

ПРИКАЗ

от 15.01.2025 № 123
г. Москва

О назначении комиссии по проведению проверки

В соответствии с Федеральным законом от 27.07.2010 № 210-ФЗ "Об организации предоставления государственных и муниципальных услуг" и постановлением Правительства Российской Федерации от 16.05.2011 № 373

ПРИКАЗЫВАЮ:

1. Утвердить состав комиссии по проведению проверки.
2. Установить срок проведения проверки до 31.12.2025.
3. Контроль за исполнением настоящего приказа возложить на заместителя министра.

Министр                                                    И.О. Фамилия
"""
    
    validator = OrderStructureValidator()
    result = validator.validate(test_text)
    
    print(validator.get_report(result))
