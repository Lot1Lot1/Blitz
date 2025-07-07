from pathlib import Path
import pandas as pd

# Папка с исходными файлами
input_dir = Path(r"D:\Python\Blitz\plots\all")

# Папки для сохранения результатов
output_dir_as = Path(r"D:\Python\Blitz\plots\as")
output_dir_dis = Path(r"D:\Python\Blitz\plots\dis")
for csv_file in input_dir.glob("*.csv"):
    df = pd.read_csv(csv_file, header=0)

    # Разделение данных
    df_part1 = df[df['Time (s)'] <= 119.6].copy()
    df_part2 = df[df['Time (s)'] >= 120].copy()

    # Перезапись времени с нуля
    df_part1['Time (s)'] = df_part1['Time (s)'] - df_part1['Time (s)'].iloc[0]
    df_part2['Time (s)'] = round(df_part2['Time (s)'] - df_part2['Time (s)'].iloc[0], 2)

    # Формируем пути для сохранения
    output_file_as = output_dir_as / f"{csv_file.stem}_as.csv"
    output_file_dis = output_dir_dis / f"{csv_file.stem}_dis.csv"

    # Сохраняем
    df_part1.to_csv(output_file_as, index=False)
    df_part2.to_csv(output_file_dis, index=False)