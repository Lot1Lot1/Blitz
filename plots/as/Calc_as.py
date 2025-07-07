import originpro as op
import pandas as pd

# Чтение данных
data = pd.read_csv("ZE15_500_Ag_as.csv")

# Фильтрация данных (первые 30 секунд)
filtered_data = data[(data['Time (s)'] >= 2) & (data['Time (s)'] <= 50)].copy()


# Загрузка отфильтрованных данных в Origin
ws = op.new_sheet()
ws.from_df(filtered_data)

# Аппроксимация ExpDecay1
fit = op.NLFit('ExpDecay1')  # Убедитесь в правильности названия функции!
fit.set_data(ws, 0, 1)  # Столбцы X (0) и Y (1)

# Установка начальных параметров
fit.parameters = {
    'y0': filtered_data.iloc[-1, 1],  # Последнее значение Y
    'A': filtered_data.iloc[0, 1] - filtered_data.iloc[-1, 1],  # Амплитуда
    't1': 5.0  # Начальное значение времени затухания
}

# Выполнение аппроксимации
fit.fit()

# Получение результатов
result = fit.result()

# ПРАВИЛЬНОЕ извлечение параметра t1 и его ошибки
t1_value = result['t1']  # Значение параметра t1
t1_error = result['e_t1']  # Стандартная ошибка параметра t1

print(f"Параметр t1: {t1_value:.4f} {t1_error:.4f}")

# Дополнительная информация о качестве аппроксимации
print(f"R-квадрат: {result['r']**2:.4f}")
print(f"Количество итераций: {result['niter']}")

op.exit()