"""
ТЕСТ ПАМЯТИ - Игра для диагностики когнитивных способностей
Автор: Система оценки памяти
Версия: 2.0

Игра предназначена для оценки кратковременной памяти и внимания.
Результаты оцениваются по трём категориям: Хорошо, Нормально, Плохо.
"""

# Импорт необходимых библиотек
import pygame  # Основная библиотека для создания игр
import sys  # Для работы с системными функциями (выход из программы)
import random  # Для случайного перемешивания карточек
import time  # Для замера времени игры
import json  # Для сохранения настроек в формате JSON
import os  # Для работы с файловой системой
from datetime import datetime  # Для записи даты и времени игры
import pandas as pd  # Для работы с Excel-файлами статистики

# ============================================================
# ИНИЦИАЛИЗАЦИЯ PYGAME (ДОЛЖНА БЫТЬ ПЕРВОЙ!)
# ============================================================
pygame.init()  # Инициализируем все модули Pygame
pygame.mixer.init()  # Инициализируем звуковую систему

# ============================================================
# НАСТРОЙКИ ИГРЫ (КОНСТАНТЫ)
# ============================================================

# Размеры окна игры (ширина, высота)
BASE_WIDTH, BASE_HEIGHT = 800, 700

# Отступы между карточками и от верхнего края
CARD_MARGIN = 10  # Расстояние между карточками в пикселях
TOP_OFFSET = 100  # Отступ сверху для счётчиков и заголовка

# ============================================================
# ЦВЕТА (RGB-формат)
# ============================================================
WHITE = (255, 255, 255)        # Белый - фон
BLACK = (0, 0, 0)              # Чёрный - текст и рамки
GRAY = (200, 200, 200)         # Серый - найденные пары
BLUE = (100, 100, 255)         # Синий - рубашка карточки
GREEN = (100, 255, 100)        # Зелёный - победа и хороший результат
RED = (255, 100, 100)          # Красный - ошибки и плохой результат
YELLOW = (255, 255, 100)       # Жёлтый - кнопки при наведении
DARK_BLUE = (50, 50, 150)      # Тёмно-синий - заголовки
LIGHT_GRAY = (220, 220, 220)   # Светло-серый - фон меню
ORANGE = (255, 165, 0)         # Оранжевый - дополнительные цвета
PURPLE = (160, 32, 240)        # Фиолетовый - дополнительные цвета
CYAN = (0, 255, 255)           # Бирюзовый - дополнительные цвета
PINK = (255, 192, 203)         # Розовый - дополнительные цвета

# ============================================================
# ШРИФТЫ (СОЗДАЁМ ПОСЛЕ ИНИЦИАЛИЗАЦИИ PYGAME)
# ============================================================
# Используем стандартные шрифты pygame
font = pygame.font.Font(None, 36)        # Обычный шрифт (36 пикселей)
big_font = pygame.font.Font(None, 48)    # Большой шрифт (48 пикселей)
title_font = pygame.font.Font(None, 72)  # Заголовочный шрифт (72 пикселя)

# ============================================================
# ПУТИ К ФАЙЛАМ
# ============================================================
SETTINGS_FILE = "memo_settings.json"  # Файл для сохранения настроек
STATS_FILE = "game_stats.xlsx"         # Файл для сохранения статистики в Excel

# ============================================================
# КЛАСС ДЛЯ УПРАВЛЕНИЯ СТАТИСТИКОЙ
# ============================================================
class StatsManager:
    """
    Класс отвечает за сохранение и загрузку статистики игр.
    Статистика сохраняется в Excel-файл для последующего анализа.
    """

    def __init__(self):
        """Инициализация менеджера статистики"""
        self.stats_file = STATS_FILE
        self.ensure_file_exists()  # Проверяем, существует ли файл

    def ensure_file_exists(self):
        """
        Проверяет наличие файла статистики.
        Если файла нет - создаёт новый с правильными колонками.
        """
        if not os.path.exists(self.stats_file):
            # Создаём пустой DataFrame с нужными колонками
            df = pd.DataFrame(columns=[
                'Дата', 'Время',           # Когда была игра
                'Режим_тип', 'Режим_сетка', # Что за режим
                'Ходы', 'Ошибки',           # Показатели игры
                'Время_игры_сек',           # Время в секундах
                'Результат', 'Оценка'       # Итог и оценка
            ])
            df.to_excel(self.stats_file, index=False)  # Сохраняем без индексов

    def save_game_result(self, game_type, grid_size, attempts, mistakes, game_time, result, rating):
        """
        Сохраняет результат игры в Excel-файл.

        Параметры:
        - game_type: тип игры (numbers/colors/shapes)
        - grid_size: размер сетки (например, "4x4")
        - attempts: количество ходов
        - mistakes: количество ошибок
        - game_time: время игры в секундах
        - result: результат (Победа/Поражение)
        - rating: оценка (Хорошо/Нормально/Плохо)
        """
        now = datetime.now()  # Текущая дата и время

        # Создаём строку с результатами
        df_new = pd.DataFrame([{
            'Дата': now.strftime('%Y-%m-%d'),      # Дата в формате ГГГГ-ММ-ДД
            'Время': now.strftime('%H:%M:%S'),     # Время в формате ЧЧ:ММ:СС
            'Режим_тип': game_type,                # Тип теста
            'Режим_сетка': grid_size,              # Размер сетки
            'Ходы': attempts,                      # Количество ходов
            'Ошибки': mistakes,                    # Количество ошибок
            'Время_игры_сек': round(game_time, 2), # Время с округлением
            'Результат': result,                   # Результат
            'Оценка': rating                       # Качественная оценка
        }])

        # Добавляем к существующему файлу или создаём новый
        if os.path.exists(self.stats_file):
            try:
                df_existing = pd.read_excel(self.stats_file)
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                df_combined.to_excel(self.stats_file, index=False)
            except Exception as e:
                print(f"Ошибка при чтении файла статистики: {e}")
                df_new.to_excel(self.stats_file, index=False)
        else:
            df_new.to_excel(self.stats_file, index=False)

        print(f"Результат сохранён в {self.stats_file}")

    def get_statistics(self):
        """
        Возвращает сводную статистику для отображения в меню.

        Возвращает словарь с:
        - total_games: всего игр
        - wins: побед
        - good_count: хороших результатов
        - normal_count: нормальных результатов
        - bad_count: плохих результатов
        - avg_attempts: среднее число ходов
        - avg_time: среднее время
        - best_time: лучшее время
        """
        if os.path.exists(self.stats_file):
            try:
                df = pd.read_excel(self.stats_file)
                if len(df) == 0:
                    return None

                # Подсчитываем различные показатели
                total_games = len(df)
                wins = len(df[df['Результат'] == 'Победа'])
                good_count = len(df[df['Оценка'] == 'Хорошо'])
                normal_count = len(df[df['Оценка'] == 'Нормально'])
                bad_count = len(df[df['Оценка'] == 'Плохо'])

                # Вычисляем средние значения
                avg_attempts = df['Ходы'].mean()
                avg_time = df['Время_игры_сек'].mean()
                best_time = df['Время_игры_сек'].min()

                return {
                    'total_games': total_games,
                    'wins': wins,
                    'win_rate': (wins / total_games * 100) if total_games > 0 else 0,
                    'good_count': good_count,
                    'normal_count': normal_count,
                    'bad_count': bad_count,
                    'avg_attempts': round(avg_attempts, 1),
                    'avg_time': round(avg_time, 1),
                    'best_time': round(best_time, 1)
                }
            except Exception as e:
                print(f"Ошибка при чтении статистики: {e}")
                return None
        return None

# Создаём глобальный экземпляр менеджера статистики
stats_manager = StatsManager()

# ============================================================
# КЛАССЫ ДЛЯ УПРАВЛЕНИЯ ЗВУКОМ И МУЗЫКОЙ
# ============================================================

class SettingsManager:
    """
    Класс для управления настройками игры:
    - Громкость музыки и звуков
    - Включение/выключение звука
    - Выбор фоновой мелодии
    """

    def __init__(self):
        """Инициализация настроек со значениями по умолчанию"""
        self.music_volume = 0.5      # Громкость музыки (0-1)
        self.sfx_volume = 0.7        # Громкость звуков (0-1)
        self.music_enabled = True     # Включена ли музыка
        self.sfx_enabled = True       # Включены ли звуки
        self.current_track = "main_theme"  # Текущая мелодия
        self.available_tracks = ["main_theme", "relax", "puzzle"]  # Доступные мелодии
        self.load()  # Загружаем сохранённые настройки

    def load(self):
        """Загружает настройки из JSON-файла"""
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    data = json.load(f)
                    self.music_volume = data.get('music_volume', 0.5)
                    self.sfx_volume = data.get('sfx_volume', 0.7)
                    self.music_enabled = data.get('music_enabled', True)
                    self.sfx_enabled = data.get('sfx_enabled', True)
                    self.current_track = data.get('current_track', 'main_theme')
            except:
                pass  # Если файл повреждён, оставляем значения по умолчанию

    def save(self):
        """Сохраняет текущие настройки в JSON-файл"""
        data = {
            'music_volume': self.music_volume,
            'sfx_volume': self.sfx_volume,
            'music_enabled': self.music_enabled,
            'sfx_enabled': self.sfx_enabled,
            'current_track': self.current_track
        }
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(data, f)


class SoundManager:
    """
    Класс для управления звуковыми эффектами:
    - Переворот карточки
    - Совпадение пары
    - Ошибка
    - Победа
    - Нажатие на кнопку
    """

    def __init__(self, settings):
        self.settings = settings
        self.sounds = {}  # Словарь для хранения звуков
        self.load_sounds()  # Загружаем звуковые файлы

    def load_sounds(self):
        """Загружает звуковые файлы из папки assets/sounds/"""
        sound_files = {
            'flip': 'assets/sounds/flip.wav',      # Звук переворота
            'match': 'assets/sounds/match.wav',    # Звук совпадения
            'mismatch': 'assets/sounds/mismatch.wav', # Звук ошибки
            'win': 'assets/sounds/win.wav',        # Звук победы
            'click': 'assets/sounds/click.wav'     # Звук кнопки
        }
        for name, path in sound_files.items():
            if os.path.exists(path):
                self.sounds[name] = pygame.mixer.Sound(path)
            else:
                self.sounds[name] = None  # Если файла нет, создаём заглушку

    def play_sound(self, name):
        """
        Воспроизводит звук по имени.
        Звук воспроизводится только если звук включён в настройках.
        """
        if not self.settings.sfx_enabled:
            return
        sound = self.sounds.get(name)
        if sound:
            sound.set_volume(self.settings.sfx_volume)
            sound.play()

    def set_sfx_volume(self, volume):
        """Устанавливает громкость всех звуковых эффектов"""
        self.settings.sfx_volume = volume
        for sound in self.sounds.values():
            if sound:
                sound.set_volume(volume)


class MusicManager:
    """
    Класс для управления фоновой музыкой:
    - Воспроизведение мелодий
    - Регулировка громкости
    - Включение/выключение
    """

    def __init__(self, settings):
        self.settings = settings
        self.current_track = None  # Текущая воспроизводимая мелодия

    def play_music(self, track_name):
        """
        Начинает воспроизведение фоновой музыки.
        Поддерживает форматы MP3 и OGG.
        """
        if not self.settings.music_enabled:
            return

        # Пробуем загрузить MP3 или OGG
        track_file = f"assets/music/{track_name}.mp3"
        if not os.path.exists(track_file):
            track_file = f"assets/music/{track_name}.ogg"

        # Воспроизводим только если файл существует и это не текущая мелодия
        if os.path.exists(track_file) and track_name != self.current_track:
            pygame.mixer.music.load(track_file)
            pygame.mixer.music.set_volume(self.settings.music_volume)
            pygame.mixer.music.play(-1)  # -1 означает бесконечное зацикливание
            self.current_track = track_name

    def stop_music(self):
        """Останавливает воспроизведение музыки"""
        pygame.mixer.music.stop()
        self.current_track = None

    def set_volume(self, volume):
        """Устанавливает громкость музыки"""
        self.settings.music_volume = volume
        if self.settings.music_enabled:
            pygame.mixer.music.set_volume(volume)

    def set_enabled(self, enabled):
        """
        Включает или выключает музыку.
        При включении восстанавливает воспроизведение с последней позиции.
        """
        self.settings.music_enabled = enabled
        if enabled:
            if self.current_track:
                pygame.mixer.music.set_volume(self.settings.music_volume)
                pygame.mixer.music.unpause()  # Возобновляем с паузы
            else:
                self.play_music(self.settings.current_track)
        else:
            pygame.mixer.music.pause()  # Ставим на паузу

# Создаём глобальные экземпляры менеджеров
settings = SettingsManager()
sound_mgr = SoundManager(settings)
music_mgr = MusicManager(settings)

# ============================================================
# КЛАСС КАРТОЧКИ
# ============================================================

class Card:
    """
    Класс представляет одну карточку в игре.
    Карточка может быть разных типов: число, цвет, форма.
    """

    def __init__(self, pair_id, x, y, size, card_type="number", color=None, shape=None):
        """
        Инициализация карточки.

        Параметры:
        - pair_id: уникальный идентификатор пары (у парных карточек одинаковый)
        - x, y: координаты левого верхнего угла
        - size: размер карточки (ширина и высота)
        - card_type: тип карточки (number/color/shape)
        - color: цвет (для режима цветов и форм)
        - shape: форма (для режима форм)
        """
        self.pair_id = pair_id          # ID пары для сравнения
        self.x = x                      # Координата X
        self.y = y                      # Координата Y
        self.size = size                # Размер
        self.card_type = card_type      # Тип отображения
        self.color = color              # Цвет для отрисовки
        self.shape = shape              # Форма для отрисовки
        self.is_flipped = False         # Перевёрнута ли карточка
        self.is_matched = False         # Найдена ли пара
        self.rect = pygame.Rect(x, y, size, size)  # Прямоугольник для кликов

    def draw(self, surface):
        """
        Отрисовывает карточку на экране.

        Состояния карточки:
        1. Найдена (is_matched = True) - серая с галочкой
        2. Закрыта (is_flipped = False) - синяя с вопросом
        3. Открыта (is_flipped = True) - белая с содержимым
        """
        # Состояние 1: карточка уже найдена (пара собрана)
        if self.is_matched:
            pygame.draw.rect(surface, GRAY, self.rect)           # Серый фон
            pygame.draw.rect(surface, BLACK, self.rect, 2)       # Чёрная рамка
            text = font.render("✓", True, GREEN)                 # Зелёная галочка
            text_rect = text.get_rect(center=self.rect.center)
            surface.blit(text, text_rect)
            return

        # Состояние 2: карточка закрыта (рубашка)
        if not self.is_flipped:
            pygame.draw.rect(surface, BLUE, self.rect)           # Синий фон
            pygame.draw.rect(surface, BLACK, self.rect, 3)       # Чёрная рамка
            text = font.render("?", True, WHITE)                 # Белый вопрос
            text_rect = text.get_rect(center=self.rect.center)
            surface.blit(text, text_rect)
            return

        # Состояние 3: карточка открыта (показываем содержимое)
        pygame.draw.rect(surface, WHITE, self.rect)              # Белый фон
        pygame.draw.rect(surface, BLACK, self.rect, 3)           # Чёрная рамка

        # Рисуем содержимое в зависимости от типа карточки
        if self.card_type == "number":
            # Режим ЧИСЛА: показываем цифру
            text = font.render(str(self.pair_id), True, BLACK)
            text_rect = text.get_rect(center=self.rect.center)
            surface.blit(text, text_rect)

        elif self.card_type == "color":
            # Режим ЦВЕТА: показываем цветной прямоугольник
            color_rect = self.rect.inflate(-20, -20)  # Уменьшаем на 20 пикселей с каждой стороны
            pygame.draw.rect(surface, self.color, color_rect)
            pygame.draw.rect(surface, BLACK, color_rect, 2)

        elif self.card_type == "shape":
            # Режим ФОРМЫ: рисуем геометрическую фигуру
            center = self.rect.center
            r = self.size // 4  # Радиус/размер фигуры

            if self.shape == "circle":
                # Круг
                pygame.draw.circle(surface, self.color, center, r)
            elif self.shape == "square":
                # Квадрат
                square_rect = pygame.Rect(center[0]-r, center[1]-r, r*2, r*2)
                pygame.draw.rect(surface, self.color, square_rect)
            elif self.shape == "triangle":
                # Треугольник
                points = [
                    (center[0], center[1]-r),           # Верхняя точка
                    (center[0]-r, center[1]+r),         # Левая точка
                    (center[0]+r, center[1]+r)          # Правая точка
                ]
                pygame.draw.polygon(surface, self.color, points)
            elif self.shape == "diamond":
                # Ромб
                points = [
                    (center[0], center[1]-r),           # Верх
                    (center[0]+r, center[1]),           # Право
                    (center[0], center[1]+r),           # Низ
                    (center[0]-r, center[1])            # Лево
                ]
                pygame.draw.polygon(surface, self.color, points)

    def handle_click(self, pos):
        """
        Проверяет, попал ли клик в область карточки.
        Карточка реагирует на клик, если она не найдена и не открыта.
        """
        return self.rect.collidepoint(pos) and not self.is_matched and not self.is_flipped

# ============================================================
# ФУНКЦИИ СОЗДАНИЯ КОЛОДЫ КАРТОЧЕК
# ============================================================

def create_board_general(rows, cols, card_size, game_mode):
    """
    Создаёт перемешанную колоду карточек для выбранного режима.

    Параметры:
    - rows: количество строк в сетке
    - cols: количество столбцов в сетке
    - card_size: размер карточки (не используется напрямую, нужен для совместимости)
    - game_mode: тип игры (numbers/colors/shapes)

    Возвращает: список словарей с данными карточек
    """
    total_cards = rows * cols
    num_pairs = total_cards // 2  # Количество пар

    if game_mode == "numbers":
        # Режим ЧИСЛА: простые числа от 1 до num_pairs
        values = list(range(1, num_pairs + 1)) * 2  # Каждое число повторяется дважды
        random.shuffle(values)  # Перемешиваем
        cards_data = [{"pair_id": v, "card_type": "number", "color": None, "shape": None} for v in values]

    elif game_mode == "colors":
        # Режим ЦВЕТА: яркие цвета
        color_palette = [RED, GREEN, BLUE, YELLOW, PURPLE, CYAN, ORANGE, PINK,
                         (128,0,128), (0,128,128), (255,140,0), (0,255,128)]
        pairs = []
        for i in range(num_pairs):
            # Берём цвета по кругу (если пар больше, чем цветов)
            color = color_palette[i % len(color_palette)]
            pairs.append({"pair_id": i, "card_type": "color", "color": color, "shape": None})
        cards_data = pairs + pairs  # Дублируем для пар
        random.shuffle(cards_data)  # Перемешиваем

    elif game_mode == "shapes":
        # Режим ФОРМЫ: комбинация цвета и геометрической фигуры
        shapes_list = ["circle", "square", "triangle", "diamond"]
        color_palette = [RED, GREEN, BLUE, YELLOW, PURPLE, CYAN, ORANGE, PINK]

        # Создаём все возможные комбинации цветов и форм
        combinations = []
        for color in color_palette:
            for shape in shapes_list:
                combinations.append((color, shape))

        # Выбираем нужное количество уникальных комбинаций
        selected = combinations[:num_pairs]
        pairs = []
        for i, (color, shape) in enumerate(selected):
            pairs.append({"pair_id": i, "card_type": "shape", "color": color, "shape": shape})
        cards_data = pairs + pairs  # Дублируем для пар
        random.shuffle(cards_data)  # Перемешиваем

    return cards_data


def create_cards_from_data(cards_data, rows, cols, card_size):
    """
    Преобразует данные карточек в объекты Card и размещает их на игровом поле.

    Параметры:
    - cards_data: список словарей с данными
    - rows: количество строк
    - cols: количество столбцов
    - card_size: размер карточки

    Возвращает: список объектов Card
    """
    # Вычисляем общую ширину сетки
    total_width = cols * card_size + (cols - 1) * CARD_MARGIN
    # Вычисляем координату X для центрирования сетки
    start_x = (BASE_WIDTH - total_width) // 2
    # Координата Y для верхнего ряда (с отступом сверху)
    start_y = TOP_OFFSET

    cards = []
    idx = 0  # Индекс текущей карточки в списке данных

    # Создаём карточки для каждой ячейки сетки
    for i in range(rows):          # Проходим по строкам
        for j in range(cols):      # Проходим по столбцам
            # Вычисляем координаты карточки
            x = start_x + j * (card_size + CARD_MARGIN)
            y = start_y + i * (card_size + CARD_MARGIN)

            # Берём данные для текущей карточки
            data = cards_data[idx]

            # Создаём объект Card
            card = Card(
                pair_id=data["pair_id"],
                x=x, y=y, size=card_size,
                card_type=data["card_type"],
                color=data["color"],
                shape=data["shape"]
            )
            cards.append(card)
            idx += 1  # Переходим к следующей карточке

    return cards

# ============================================================
# ФУНКЦИЯ ОЦЕНКИ РЕЗУЛЬТАТА
# ============================================================

def evaluate_performance(attempts, mistakes, total_pairs, game_time):
    """
    Оценивает результат игры по трём категориям: Хорошо, Нормально, Плохо.

    Критерии оценки:
    - Количество ходов (чем меньше, тем лучше)
    - Количество ошибок (чем меньше, тем лучше)
    - Время игры (чем быстрее, тем лучше)

    Алгоритм:
    - Теоретический минимум ходов = количество пар
    - Ожидаемое количество ходов для здорового человека = количество пар * 1.5
    - Ожидаемое количество ошибок = количество пар * 0.5
    - Ожидаемое время = количество пар * 5 секунд

    Возвращает: (оценка, цвет_оценки)
    """
    # Теоретический минимум ходов = количество пар
    min_attempts = total_pairs
    # Ожидаемое количество ходов для здорового человека
    expected_attempts = total_pairs * 1.5
    # Ожидаемое количество ошибок
    expected_mistakes = total_pairs * 0.5

    score = 0  # Суммарный балл (макс 5, мин -1)

    # ОЦЕНКА ПО ХОДАМ (макс 2 балла)
    if attempts <= expected_attempts:
        score += 2  # Отлично: уложились в норму
    elif attempts <= expected_attempts * 1.5:
        score += 1  # Хорошо: немного превысили норму
    else:
        score += 0  # Плохо: сильно превысили норму

    # ОЦЕНКА ПО ОШИБКАМ (макс 2 балла)
    if mistakes <= expected_mistakes:
        score += 2  # Отлично: мало ошибок
    elif mistakes <= expected_mistakes * 2:
        score += 1  # Хорошо: терпимое количество ошибок
    else:
        score += 0  # Плохо: много ошибок

    # ОЦЕНКА ПО ВРЕМЕНИ (макс 1 балл, может быть штраф)
    expected_time = total_pairs * 5  # База: 5 секунд на пару
    if game_time <= expected_time:
        score += 1  # Бонус за быстроту
    elif game_time > expected_time * 2:
        score -= 1  # Штраф за медлительность

    # ИТОГОВАЯ ОЦЕНКА
    if score >= 4:
        return "Хорошо", GREEN      # Отличный результат
    elif score >= 2:
        return "Нормально", YELLOW  # Результат в пределах нормы
    else:
        return "Плохо", RED         # Результат ниже нормы

# ============================================================
# ФУНКЦИИ ОТРИСОВКИ ИГРОВОГО ИНТЕРФЕЙСА
# ============================================================

def draw_game_state(screen, attempts, matched_pairs, total_pairs, game_over, mode_str, mistakes):
    """
    Отрисовывает всю информацию об игре на экране:
    - Счётчик ходов
    - Счётчик ошибок
    - Количество найденных пар
    - Название режима
    """
    # Отображаем количество ходов в левом верхнем углу
    attempts_text = font.render(f"Ходы: {attempts}", True, BLACK)
    screen.blit(attempts_text, (20, 20))

    # Отображаем количество ошибок (красным цветом)
    mistakes_text = font.render(f"Ошибки: {mistakes}", True, RED if mistakes > 0 else BLACK)
    screen.blit(mistakes_text, (20, 60))

    # Отображаем прогресс по парам в правом верхнем углу
    pairs_text = font.render(f"Пары: {matched_pairs} / {total_pairs}", True, BLACK)
    screen.blit(pairs_text, (BASE_WIDTH - 200, 20))

    # Отображаем название режима по центру вверху
    mode_text = font.render(mode_str, True, DARK_BLUE)
    screen.blit(mode_text, (BASE_WIDTH // 2 - mode_text.get_width()//2, 20))


def show_result_screen(screen, attempts, mistakes, total_pairs, game_time, game_type, grid_size):
    """
    Показывает экран с результатами после окончания игры.
    Отображает:
    - Статистику игры
    - Оценку (Хорошо/Нормально/Плохо)
    - Рекомендации
    - Кнопки для продолжения
    """
    # Получаем оценку и её цвет
    rating, rating_color = evaluate_performance(attempts, mistakes, total_pairs, game_time)

    # Сохраняем результат в Excel
    grid_size_str = f"{grid_size[0]}x{grid_size[1]}"
    stats_manager.save_game_result(
        game_type=game_type,
        grid_size=grid_size_str,
        attempts=attempts,
        mistakes=mistakes,
        game_time=game_time,
        result="Победа",
        rating=rating
    )

    waiting_for_input = True
    clock = pygame.time.Clock()

    while waiting_for_input:
        screen.fill(WHITE)  # Очищаем экран

        # ЗАГОЛОВОК
        title_text = big_font.render("РЕЗУЛЬТАТ ТЕСТИРОВАНИЯ", True, DARK_BLUE)
        title_rect = title_text.get_rect(center=(BASE_WIDTH // 2, 100))
        screen.blit(title_text, title_rect)

        # СТАТИСТИКА ИГРЫ
        stats_y = 200
        stats = [
            f"Режим: {game_type.capitalize()} {grid_size_str}",
            f"Ходов сделано: {attempts}",
            f"Ошибок допущено: {mistakes}",
            f"Время прохождения: {game_time:.1f} секунд"
        ]

        for stat in stats:
            stat_text = font.render(stat, True, BLACK)
            stat_rect = stat_text.get_rect(center=(BASE_WIDTH // 2, stats_y))
            screen.blit(stat_text, stat_rect)
            stats_y += 50

        # ОЦЕНКА (крупным шрифтом)
        rating_y = 380
        rating_text = big_font.render(f"ОЦЕНКА: {rating}", True, rating_color)
        rating_rect = rating_text.get_rect(center=(BASE_WIDTH // 2, rating_y))
        screen.blit(rating_text, rating_rect)

        # РЕКОМЕНДАЦИЯ в зависимости от оценки
        explanation_y = 450
        if rating == "Хорошо":
            explanation = "Память в отличном состоянии! Так держать!"
            explanation_color = GREEN
        elif rating == "Нормально":
            explanation = "Результат в пределах нормы. Рекомендуются тренировки памяти."
            explanation_color = YELLOW
        else:
            explanation = "Результат ниже нормы. Рекомендуется консультация специалиста."
            explanation_color = RED

        explanation_text = font.render(explanation, True, explanation_color)
        explanation_rect = explanation_text.get_rect(center=(BASE_WIDTH // 2, explanation_y))
        screen.blit(explanation_text, explanation_rect)

        # КНОПКИ
        button_width = 200
        button_height = 50
        spacing = 30
        start_x = (BASE_WIDTH - (button_width * 2 + spacing)) // 2
        buttons_y = BASE_HEIGHT - 100

        # Кнопка "Новая игра"
        new_game_rect = pygame.Rect(start_x, buttons_y, button_width, button_height)
        pygame.draw.rect(screen, GREEN, new_game_rect)
        pygame.draw.rect(screen, BLACK, new_game_rect, 3)
        new_game_text = font.render("Новая игра", True, BLACK)
        new_game_rect_text = new_game_text.get_rect(center=new_game_rect.center)
        screen.blit(new_game_text, new_game_rect_text)

        # Кнопка "Главное меню"
        menu_rect = pygame.Rect(start_x + button_width + spacing, buttons_y, button_width, button_height)
        pygame.draw.rect(screen, BLUE, menu_rect)
        pygame.draw.rect(screen, BLACK, menu_rect, 3)
        menu_text = font.render("Главное меню", True, WHITE)
        menu_rect_text = menu_text.get_rect(center=menu_rect.center)
        screen.blit(menu_text, menu_rect_text)

        pygame.display.flip()  # Обновляем экран

        # ОБРАБОТКА СОБЫТИЙ
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()
            if event.type == pygame.MOUSEBUTTONDOWN:
                if new_game_rect.collidepoint(event.pos):
                    return "restart"  # Перезапустить игру
                if menu_rect.collidepoint(event.pos):
                    return "menu"     # Вернуться в меню
            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_ESCAPE:
                    return "menu"     # ESC тоже возвращает в меню

        clock.tick(60)  # Ограничиваем FPS

# ============================================================
# ОСНОВНОЙ ИГРОВОЙ ЦИКЛ
# ============================================================

def game_loop(rows, cols, game_mode):
    """
    Главный игровой цикл.
    Содержит логику игры: открытие карточек, проверка пар, обработка ошибок.

    Параметры:
    - rows: количество строк в сетке
    - cols: количество столбцов в сетке
    - game_mode: тип игры (numbers/colors/shapes)
    """
    # Вычисляем оптимальный размер карточки, чтобы всё поместилось
    max_card_width = (BASE_WIDTH - (cols - 1) * CARD_MARGIN) // cols
    max_card_height = (BASE_HEIGHT - TOP_OFFSET - (rows - 1) * CARD_MARGIN - 50) // rows
    card_size = min(max_card_width, max_card_height, 100)  # Не больше 100 пикселей

    # Создаём колоду карточек
    cards_data = create_board_general(rows, cols, card_size, game_mode)
    cards = create_cards_from_data(cards_data, rows, cols, card_size)
    total_pairs = len(cards) // 2  # Общее количество пар

    # Игровые переменные
    attempts = 0        # Количество ходов
    mistakes = 0        # Количество ошибок
    matched_pairs = 0   # Количество найденных пар
    game_over = False   # Завершена ли игра

    # Засекаем время начала игры
    start_time = time.time()
    game_time = 0

    # Переменные для ожидания между двумя открытыми карточками
    waiting = False           # Флаг ожидания
    wait_start_time = 0       # Время начала ожидания
    first_card = None         # Первая открытая карточка
    second_card = None        # Вторая открытая карточка

    clock = pygame.time.Clock()
    running = True
    result = "menu"

    # Запускаем фоновую музыку
    music_mgr.play_music(settings.current_track)

    while running:
        screen = pygame.display.get_surface()
        screen.fill(WHITE)  # Очищаем экран
        current_time = time.time()

        # ===== ОБРАБОТКА СОБЫТИЙ =====
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_ESCAPE:
                    running = False
                    result = "menu"

            if event.type == pygame.MOUSEBUTTONDOWN and not game_over and not waiting:
                pos = pygame.mouse.get_pos()
                for card in cards:
                    if card.handle_click(pos):
                        # Воспроизводим звук переворота
                        sound_mgr.play_sound('flip')
                        card.is_flipped = True  # Открываем карточку

                        if first_card is None:
                            # Это первая открытая карточка
                            first_card = card
                        elif second_card is None and card != first_card:
                            # Это вторая открытая карточка
                            second_card = card
                            attempts += 1  # Увеличиваем счётчик ходов
                            waiting = True  # Начинаем ожидание
                            wait_start_time = current_time
                        break

        # ===== ПРОВЕРКА ПАРЫ ПОСЛЕ ЗАДЕРЖКИ =====
        if waiting and current_time - wait_start_time > 0.8:  # Ждём 0.8 секунды
            waiting = False
            if first_card and second_card:
                if first_card.pair_id == second_card.pair_id:
                    # ПАРА НАЙДЕНА!
                    sound_mgr.play_sound('match')
                    first_card.is_matched = True
                    second_card.is_matched = True
                    matched_pairs += 1

                    # Проверяем, собраны ли все пары
                    if matched_pairs == total_pairs:
                        game_over = True
                        game_time = time.time() - start_time
                        sound_mgr.play_sound('win')

                        # Показываем экран с результатами
                        action = show_result_screen(
                            screen, attempts, mistakes, total_pairs,
                            game_time, game_mode, (rows, cols)
                        )

                        if action == "restart":
                            # Перезапуск игры
                            cards_data = create_board_general(rows, cols, card_size, game_mode)
                            cards = create_cards_from_data(cards_data, rows, cols, card_size)
                            attempts = 0
                            mistakes = 0
                            matched_pairs = 0
                            game_over = False
                            start_time = time.time()
                            waiting = False
                            first_card = None
                            second_card = None
                        else:
                            running = False
                            result = "menu"
                else:
                    # ПАРА НЕ СОВПАЛА - ОШИБКА!
                    sound_mgr.play_sound('mismatch')
                    mistakes += 1  # Увеличиваем счётчик ошибок
                    # Переворачиваем карточки обратно
                    first_card.is_flipped = False
                    second_card.is_flipped = False

                # Сбрасываем выбранные карточки
                first_card = None
                second_card = None

        # ===== ОТРИСОВКА =====
        for card in cards:
            card.draw(screen)  # Рисуем все карточки

        # Отображаем игровую статистику
        mode_display = f"{game_mode.capitalize()} {rows}x{cols}"
        draw_game_state(screen, attempts, matched_pairs, total_pairs, game_over, mode_display, mistakes)

        pygame.display.flip()  # Обновляем экран
        clock.tick(60)  # Ограничиваем FPS до 60

    return result

# ============================================================
# МЕНЮ СТАТИСТИКИ
# ============================================================

def statistics_menu(screen):
    """
    Отображает окно со сводной статистикой всех игр.
    Показывает общее количество игр, побед, оценок и средние показатели.
    """
    clock = pygame.time.Clock()
    running = True
    stats = stats_manager.get_statistics()  # Получаем статистику

    while running:
        screen.fill(LIGHT_GRAY)

        # Заголовок
        title_text = big_font.render("Статистика тестирований", True, DARK_BLUE)
        title_rect = title_text.get_rect(center=(BASE_WIDTH//2, 50))
        screen.blit(title_text, title_rect)

        if stats and stats['total_games'] > 0:
            # Формируем строки статистики
            stats_lines = [
                f"Всего тестов: {stats['total_games']}",
                f"",
                f"РЕЗУЛЬТАТЫ:",
                f"  Хорошо: {stats['good_count']}",
                f"  Нормально: {stats['normal_count']}",
                f"  Плохо: {stats['bad_count']}",
                f"",
                f"Среднее число ходов: {stats['avg_attempts']}",
                f"Среднее время: {stats['avg_time']} сек.",
                f"Лучшее время: {stats['best_time']} сек."
            ]

            # Отображаем каждую строку по центру
            y_offset = 180
            for line in stats_lines:
                text = font.render(line, True, BLACK)
                text_rect = text.get_rect(center=(BASE_WIDTH//2, y_offset))
                screen.blit(text, text_rect)
                y_offset += 40
        else:
            # Сообщение, если статистики ещё нет
            no_stats_text = font.render("Нет сохранённых результатов. Пройдите тест!", True, BLACK)
            no_stats_rect = no_stats_text.get_rect(center=(BASE_WIDTH//2, BASE_HEIGHT//2))
            screen.blit(no_stats_text, no_stats_rect)

        # Кнопка возврата
        back_rect = pygame.Rect(BASE_WIDTH//2 - 100, BASE_HEIGHT - 100, 200, 50)
        pygame.draw.rect(screen, BLUE, back_rect)
        pygame.draw.rect(screen, BLACK, back_rect, 3)
        back_text = font.render("Назад", True, WHITE)
        back_rect_text = back_text.get_rect(center=back_rect.center)
        screen.blit(back_text, back_rect_text)

        pygame.display.flip()

        # Обработка событий
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()
            if event.type == pygame.MOUSEBUTTONDOWN:
                if back_rect.collidepoint(event.pos):
                    running = False
            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_ESCAPE:
                    running = False

        clock.tick(60)

# ============================================================
# МЕНЮ НАСТРОЕК
# ============================================================

def settings_menu(screen):
    """
    Окно настроек игры.
    Позволяет регулировать громкость музыки и звуков,
    включать/выключать звук, выбирать фоновую мелодию.
    """
    clock = pygame.time.Clock()
    running = True

    # Позиции элементов интерфейса (всё центрировано)
    slider_music_rect = pygame.Rect(BASE_WIDTH//2 - 150, 180, 300, 10)
    slider_sfx_rect = pygame.Rect(BASE_WIDTH//2 - 150, 280, 300, 10)
    knob_radius = 10

    # Текущие значения из настроек
    music_vol = settings.music_volume
    sfx_vol = settings.sfx_volume
    music_enabled = settings.music_enabled
    sfx_enabled = settings.sfx_enabled
    track_index = settings.available_tracks.index(settings.current_track) if settings.current_track in settings.available_tracks else 0

    # Чекбоксы
    music_checkbox_rect = pygame.Rect(BASE_WIDTH//2 - 100, 140, 20, 20)
    sfx_checkbox_rect = pygame.Rect(BASE_WIDTH//2 - 100, 240, 20, 20)

    # Кнопки переключения мелодии
    track_left_rect = pygame.Rect(BASE_WIDTH//2 - 100, 350, 40, 40)
    track_right_rect = pygame.Rect(BASE_WIDTH//2 + 60, 350, 40, 40)
    track_display_rect = pygame.Rect(BASE_WIDTH//2 - 50, 350, 100, 40)

    # Кнопка сохранения
    save_rect = pygame.Rect(BASE_WIDTH//2 - 100, BASE_HEIGHT - 100, 200, 50)

    # Флаги для перетаскивания ползунков
    dragging_music = False
    dragging_sfx = False

    while running:
        screen.fill(LIGHT_GRAY)

        # Заголовок
        title = big_font.render("Настройки", True, DARK_BLUE)
        title_rect = title.get_rect(center=(BASE_WIDTH//2, 50))
        screen.blit(title, title_rect)

        # Подписи
        music_text = font.render("Музыка:", True, BLACK)
        screen.blit(music_text, (BASE_WIDTH//2 - 200, 140))
        sfx_text = font.render("Звуки:", True, BLACK)
        screen.blit(sfx_text, (BASE_WIDTH//2 - 200, 240))

        # Чекбокс музыки
        pygame.draw.rect(screen, WHITE if music_enabled else GRAY, music_checkbox_rect)
        pygame.draw.rect(screen, BLACK, music_checkbox_rect, 2)
        if music_enabled:
            # Рисуем галочку
            pygame.draw.line(screen, GREEN, (music_checkbox_rect.left+5, music_checkbox_rect.top+5),
                             (music_checkbox_rect.right-5, music_checkbox_rect.bottom-5), 3)
            pygame.draw.line(screen, GREEN, (music_checkbox_rect.right-5, music_checkbox_rect.top+5),
                             (music_checkbox_rect.left+5, music_checkbox_rect.bottom-5), 3)

        # Чекбокс звуков
        pygame.draw.rect(screen, WHITE if sfx_enabled else GRAY, sfx_checkbox_rect)
        pygame.draw.rect(screen, BLACK, sfx_checkbox_rect, 2)
        if sfx_enabled:
            pygame.draw.line(screen, GREEN, (sfx_checkbox_rect.left+5, sfx_checkbox_rect.top+5),
                             (sfx_checkbox_rect.right-5, sfx_checkbox_rect.bottom-5), 3)
            pygame.draw.line(screen, GREEN, (sfx_checkbox_rect.right-5, sfx_checkbox_rect.top+5),
                             (sfx_checkbox_rect.left+5, sfx_checkbox_rect.bottom-5), 3)

        # Ползунок громкости музыки
        pygame.draw.rect(screen, GRAY, slider_music_rect)
        fill_width = int(music_vol * slider_music_rect.width)
        pygame.draw.rect(screen, BLUE, (slider_music_rect.x, slider_music_rect.y, fill_width, slider_music_rect.height))
        knob_x = slider_music_rect.x + music_vol * slider_music_rect.width
        pygame.draw.circle(screen, BLACK, (int(knob_x), slider_music_rect.centery), knob_radius)
        music_vol_text = font.render(f"{int(music_vol*100)}%", True, BLACK)
        screen.blit(music_vol_text, (slider_music_rect.right+10, slider_music_rect.centery-12))

        # Ползунок громкости звуков
        pygame.draw.rect(screen, GRAY, slider_sfx_rect)
        fill_width = int(sfx_vol * slider_sfx_rect.width)
        pygame.draw.rect(screen, BLUE, (slider_sfx_rect.x, slider_sfx_rect.y, fill_width, slider_sfx_rect.height))
        knob_x = slider_sfx_rect.x + sfx_vol * slider_sfx_rect.width
        pygame.draw.circle(screen, BLACK, (int(knob_x), slider_sfx_rect.centery), knob_radius)
        sfx_vol_text = font.render(f"{int(sfx_vol*100)}%", True, BLACK)
        screen.blit(sfx_vol_text, (slider_sfx_rect.right+10, slider_sfx_rect.centery-12))

        # Выбор мелодии
        track_label = font.render("Мелодия:", True, BLACK)
        screen.blit(track_label, (BASE_WIDTH//2 - 200, 360))
        pygame.draw.rect(screen, WHITE, track_display_rect)
        pygame.draw.rect(screen, BLACK, track_display_rect, 2)
        track_name = settings.available_tracks[track_index].capitalize()
        track_text = font.render(track_name, True, BLACK)
        screen.blit(track_text, (track_display_rect.x+10, track_display_rect.y+10))

        # Кнопки переключения треков
        pygame.draw.rect(screen, BLUE, track_left_rect)
        pygame.draw.rect(screen, BLUE, track_right_rect)
        left_text = font.render("<", True, WHITE)
        right_text = font.render(">", True, WHITE)
        screen.blit(left_text, (track_left_rect.x+15, track_left_rect.y+10))
        screen.blit(right_text, (track_right_rect.x+15, track_right_rect.y+10))

        # Кнопка сохранения
        pygame.draw.rect(screen, GREEN, save_rect)
        pygame.draw.rect(screen, BLACK, save_rect, 3)
        save_text = font.render("Сохранить и выйти", True, BLACK)
        save_rect_text = save_text.get_rect(center=save_rect.center)
        screen.blit(save_text, save_rect_text)

        # Обработка событий
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

            if event.type == pygame.MOUSEBUTTONDOWN:
                if event.button == 1:
                    # Клик по чекбоксам
                    if music_checkbox_rect.collidepoint(event.pos):
                        music_enabled = not music_enabled
                        music_mgr.set_enabled(music_enabled)
                    if sfx_checkbox_rect.collidepoint(event.pos):
                        sfx_enabled = not sfx_enabled
                        settings.sfx_enabled = sfx_enabled

                    # Начало перетаскивания ползунков
                    if slider_music_rect.collidepoint(event.pos):
                        dragging_music = True
                    if slider_sfx_rect.collidepoint(event.pos):
                        dragging_sfx = True

                    # Переключение треков
                    if track_left_rect.collidepoint(event.pos):
                        track_index = (track_index - 1) % len(settings.available_tracks)
                        settings.current_track = settings.available_tracks[track_index]
                        if music_enabled:
                            music_mgr.play_music(settings.current_track)
                    if track_right_rect.collidepoint(event.pos):
                        track_index = (track_index + 1) % len(settings.available_tracks)
                        settings.current_track = settings.available_tracks[track_index]
                        if music_enabled:
                            music_mgr.play_music(settings.current_track)

                    # Сохранение настроек
                    if save_rect.collidepoint(event.pos):
                        settings.music_volume = music_vol
                        settings.sfx_volume = sfx_vol
                        settings.music_enabled = music_enabled
                        settings.sfx_enabled = sfx_enabled
                        settings.current_track = settings.available_tracks[track_index]
                        settings.save()
                        music_mgr.set_volume(music_vol)
                        sound_mgr.set_sfx_volume(sfx_vol)
                        music_mgr.set_enabled(music_enabled)
                        running = False

            if event.type == pygame.MOUSEBUTTONUP:
                if event.button == 1:
                    dragging_music = False
                    dragging_sfx = False

            if event.type == pygame.MOUSEMOTION:
                # Перетаскивание ползунка музыки
                if dragging_music:
                    rel_x = max(0, min(event.pos[0] - slider_music_rect.x, slider_music_rect.width))
                    music_vol = rel_x / slider_music_rect.width
                    if music_enabled:
                        music_mgr.set_volume(music_vol)
                # Перетаскивание ползунка звуков
                if dragging_sfx:
                    rel_x = max(0, min(event.pos[0] - slider_sfx_rect.x, slider_sfx_rect.width))
                    sfx_vol = rel_x / slider_sfx_rect.width
                    sound_mgr.set_sfx_volume(sfx_vol)

        pygame.display.flip()
        clock.tick(60)

# ============================================================
# ГЛАВНОЕ МЕНЮ
# ============================================================

def main_menu():
    """
    Главное меню игры.
    Состоит из двух этапов:
    1. Выбор типа теста (Числа/Цвета/Формы)
    2. Выбор уровня сложности (размер сетки)
    """
    screen = pygame.display.set_mode((BASE_WIDTH, BASE_HEIGHT))
    pygame.display.set_caption("Тест памяти - Мемори")
    clock = pygame.time.Clock()

    # Доступные типы тестов
    game_types = [
        ("Числа", "numbers"),
        ("Цвета", "colors"),
        ("Формы", "shapes")
    ]

    # Доступные размеры сетки (уровни сложности)
    size_modes = [
        ("2x2 (2 пары)", 2, 2),      # Очень легко
        ("4x4 (8 пар)", 4, 4),        # Легко
        ("6x6 (18 пар)", 6, 6),       # Средне
        ("8x4 (16 пар)", 4, 8),       # Сложно
    ]

    # Состояние меню: выбор типа или выбор размера
    state = "select_type"
    selected_game_type = None
    buttons = []
    button_height = 60
    button_width = 300
    spacing = 20

    # Запускаем фоновую музыку
    music_mgr.play_music(settings.current_track)

    running = True
    while running:
        screen.fill(LIGHT_GRAY)

        if state == "select_type":
            # ЭТАП 1: ВЫБОР ТИПА ТЕСТА
            title_text = title_font.render("ТЕСТ ПАМЯТИ", True, DARK_BLUE)
            title_rect = title_text.get_rect(center=(BASE_WIDTH // 2, 60))
            screen.blit(title_text, title_rect)

            subtitle_text = font.render("Выберите тип теста:", True, BLACK)
            subtitle_rect = subtitle_text.get_rect(center=(BASE_WIDTH // 2, 130))
            screen.blit(subtitle_text, subtitle_rect)

            # Создаём кнопки типов тестов
            start_y = 200
            buttons = []
            for i, (name, type_id) in enumerate(game_types):
                rect = pygame.Rect(BASE_WIDTH//2 - button_width//2, start_y + i*(button_height+spacing), button_width, button_height)
                buttons.append((rect, name, type_id))

            # Кнопка "Статистика"
            stats_rect = pygame.Rect(BASE_WIDTH//2 - button_width//2, start_y + len(game_types)*(button_height+spacing) + 10, button_width, button_height)
            buttons.append((stats_rect, "Статистика", "stats"))

            # Кнопка "Настройки"
            settings_rect = pygame.Rect(BASE_WIDTH//2 - button_width//2, start_y + (len(game_types)+1)*(button_height+spacing) + 10, button_width, button_height)
            buttons.append((settings_rect, "Настройки", "settings"))

            # Кнопка "Выход"
            exit_rect = pygame.Rect(BASE_WIDTH//2 - button_width//2, start_y + (len(game_types)+2)*(button_height+spacing) + 10, button_width, button_height)
            buttons.append((exit_rect, "Выход", "exit"))

        elif state == "select_size":
            # ЭТАП 2: ВЫБОР УРОВНЯ СЛОЖНОСТИ
            title_text = title_font.render(f"Тест: {selected_game_type.capitalize()}", True, DARK_BLUE)
            title_rect = title_text.get_rect(center=(BASE_WIDTH // 2, 60))
            screen.blit(title_text, title_rect)

            subtitle_text = font.render("Выберите уровень сложности:", True, BLACK)
            subtitle_rect = subtitle_text.get_rect(center=(BASE_WIDTH // 2, 130))
            screen.blit(subtitle_text, subtitle_rect)

            # Создаём кнопки размеров сетки
            start_y = 200
            buttons = []
            for i, (name, rows, cols) in enumerate(size_modes):
                rect = pygame.Rect(BASE_WIDTH//2 - button_width//2, start_y + i*(button_height+spacing), button_width, button_height)
                buttons.append((rect, name, (rows, cols)))

            # Кнопка "Назад"
            back_rect = pygame.Rect(BASE_WIDTH//2 - button_width//2, start_y + len(size_modes)*(button_height+spacing) + 20, button_width, button_height)
            buttons.append((back_rect, "Назад", "back"))

        # ОТРИСОВКА КНОПОК
        mouse_pos = pygame.mouse.get_pos()
        for rect, name, action in buttons:
            # Меняем цвет при наведении мыши
            color = YELLOW if rect.collidepoint(mouse_pos) else BLUE
            pygame.draw.rect(screen, color, rect)
            pygame.draw.rect(screen, BLACK, rect, 3)
            text = font.render(name, True, WHITE if color == BLUE else BLACK)
            text_rect = text.get_rect(center=rect.center)
            screen.blit(text, text_rect)

        pygame.display.flip()

        # ОБРАБОТКА СОБЫТИЙ
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

            if event.type == pygame.MOUSEBUTTONDOWN:
                for rect, name, action in buttons:
                    if rect.collidepoint(event.pos):
                        if state == "select_type":
                            if action == "stats":
                                statistics_menu(screen)
                                buttons = []
                                break
                            elif action == "settings":
                                settings_menu(screen)
                                music_mgr.play_music(settings.current_track)
                                buttons = []
                                break
                            elif action == "exit":
                                pygame.quit()
                                sys.exit()
                            else:
                                # Выбран тип теста, переходим к выбору размера
                                selected_game_type = action
                                state = "select_size"
                                buttons = []
                                break

                        elif state == "select_size":
                            if action == "back":
                                # Возврат к выбору типа
                                state = "select_type"
                                buttons = []
                                break
                            else:
                                # Запуск игры с выбранными параметрами
                                rows, cols = action
                                result = game_loop(rows, cols, selected_game_type)
                                if result == "exit":
                                    pygame.quit()
                                    sys.exit()
                                # Возврат в главное меню после игры
                                state = "select_type"
                                buttons = []
                                music_mgr.play_music(settings.current_track)
                                break

            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_ESCAPE:
                    if state == "select_size":
                        state = "select_type"
                        buttons = []
                    else:
                        pygame.quit()
                        sys.exit()

        clock.tick(60)

# ============================================================
# ТОЧКА ВХОДА В ПРОГРАММУ
# ============================================================

if __name__ == "__main__":
    # Создаём необходимые папки для звуков и музыки
    os.makedirs("assets/sounds", exist_ok=True)
    os.makedirs("assets/music", exist_ok=True)

    # Выводим информацию о программе в консоль
    print("=" * 50)
    print("ТЕСТ ПАМЯТИ - Диагностика когнитивных способностей")
    print("=" * 50)
    print("\nРезультаты тестирования оцениваются по трём категориям:")
    print("  ХОРОШО - отличная память, результат выше нормы")
    print("  НОРМАЛЬНО - результат в пределах возрастной нормы")
    print("  ПЛОХО - результат ниже нормы, рекомендуется консультация")
    print("\nСтатистика сохраняется в файл: game_stats.xlsx")
    print("\nДля работы звуков поместите файлы в папку assets/sounds/:")
    print("  - flip.wav (переворот карточки)")
    print("  - match.wav (совпадение пары)")
    print("  - mismatch.wav (ошибка)")
    print("  - win.wav (победа)")
    print("  - click.wav (нажатие кнопки)")
    print("\nДля музыки: assets/music/main_theme.mp3, relax.mp3, puzzle.mp3")
    print("=" * 50)

    # Запускаем главное меню
    main_menu()