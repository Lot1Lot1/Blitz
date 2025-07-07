import pandas as pd
import matplotlib.pyplot as plt
import os


def adjust_data_continuity(df):
    """Корректирует данные после 120 секунд для плавного перехода"""
    # Находим значения в моменты 119.6 и 120 секунд
    value_119_6 = df.loc[df['Time (s)'].round(1) == 119.6, 'Binding (nm)'].values
    value_120 = df.loc[df['Time (s)'].round(1) == 121.6, 'Binding (nm)'].values

    if len(value_119_6) == 0 or len(value_120) == 0:
        raise ValueError("Не найдены значения для 119.6 или 120.0 секунд")

    # Вычисляем разницу для коррекции
    adjustment = value_119_6[0] - value_120[0]

    # Применяем коррекцию ко всем данным после 120 секунд
    mask = df['Time (s)'] >= 120
    df.loc[mask, 'Binding (nm)'] = round(df.loc[mask, 'Binding (nm)'] + adjustment, 9)

    return df, adjustment

def load_and_clean_csv(filepath):
    """Загружает и очищает CSV с особым форматом"""
    with open(filepath, 'r') as f:
        # Читаем заголовок
        header = f.readline().strip()

        # Читаем остальные строки, пропуская некорректные
        data = []
        for line in f:
            parts = line.strip().split(',')
            if len(parts) == 2:  # Только строки с 2 значениями
                try:
                    time = float(parts[0].strip())
                    binding = float(parts[1].strip())
                    data.append([time, binding])
                except ValueError:
                    continue

    # Создаем DataFrame
    df = pd.DataFrame(data, columns=['Time (s)', 'Binding (nm)'])
    return df


def process_data(input_file, output_file):
    try:
        # 1. Загрузка и очистка
        df = load_and_clean_csv(input_file)
        df.to_csv("debug_01_cleaned.csv", index=False)

        # 2. Фильтрация по времени (30-270 сек)
        filtered = df[(df['Time (s)'] >= 30) & (df['Time (s)'] < 270)].copy()
        filtered.to_csv("debug_02_filtered.csv", index=False)

        if len(filtered) == 0:
            raise ValueError("Нет данных в диапазоне 30-270 секунд!")

        # 3. Нормализация времени (начинаем с 0)
        filtered['Time (s)'] = round((filtered['Time (s)'] - 30.2),2 )

        # 4. Нормализация сигнала (начинаем с 0)
        first_signal = filtered['Binding (nm)'].iloc[0]
        filtered['Binding (nm)'] = round((filtered['Binding (nm)'] - first_signal), 9)
        filtered.to_csv("debug_03_normalized.csv", index=False)

        # Корректировка непрерывности
        adjusted_df, adj_value = adjust_data_continuity(filtered)
        print(f"Применена коррекция: {adj_value:.6f} нм")

        # Сохранение
        adjusted_df.to_csv(output_file, index=False)
        print(f"Данные сохранены в {output_file}")

        # 6. Рисуем
        plt.plot( adjusted_df['Time (s)'],  adjusted_df['Binding (nm)'])
        plt.xlabel('Time (s)')
        plt.ylabel('Binding (nm)')
        plt.show()

    except Exception as e:
        print(f"❌ Ошибка: {str(e)}")



input_path = r"D:\laba\blitz Install\Data\exp\ZE 19-26 Ag 23.12.24\2025-03-05_032.csv"
output_path = r"D:\Python\Blitz\plots\all\ZE26_250_Ag.csv"
process_data(input_path, output_path)

