import originpro as op
import pandas as pd
import os
import warnings

# Настройки
input_folder = '.'  # Текущая папка (можно указать другую)
output_file = 'fit_results.csv'  # Файл для сохранения результатов
time_range = (2, 40)  # Диапазон времени для анализа

# Создаем список для хранения результатов
results = []

# Получаем список CSV-файлов в папке
csv_files = [f for f in os.listdir(input_folder) if f.endswith('.csv')]

# Проверяем доступность Origin
if not op.oext:
    raise RuntimeError("Не удалось подключиться к Origin")

# Обрабатываем каждый файл
for filename in csv_files:
    try:
        print(f"\nОбработка файла: {filename}")

        # Чтение данных
        data = pd.read_csv(os.path.join(input_folder, filename))

        # Проверяем наличие нужных столбцов
        if 'Time (s)' not in data.columns or len(data.columns) < 2:
            warnings.warn(f"Файл {filename} не содержит нужных столбцов. Пропускаем.")
            continue

        y_col = data.columns[1]  # Берем второй столбец как Y

        # Фильтрация данных
        filtered_data = data[(data['Time (s)'] >= time_range[0]) &
                             (data['Time (s)'] <= time_range[1])].copy()

        if len(filtered_data) < 10:
            warnings.warn(f"Файл {filename} содержит слишком мало точек после фильтрации. Пропускаем.")
            continue

        # Загрузка данных в Origin
        op.new_book()
        ws = op.find_sheet()
        ws.from_df(filtered_data)

        # Аппроксимация ExpDecay1
        fit = op.NLFit('ExpDecay1')
        fit.set_data(ws, 0, 1)  # Столбцы X (0) и Y (1)

        # Установка начальных параметров (более надежный способ)
        y0_guess = filtered_data.iloc[-1, 1]
        A_guess = filtered_data.iloc[0, 1] - y0_guess

        # Альтернативные названия параметров (в зависимости от версии Origin)
        param_names = {
            'y0': ['y0', 'y0'],
            'A': ['A', 'A1', 'amplitude'],
            't1': ['t1', 'tau1']
        }

        # Пытаемся установить параметры разными способами
        success = False
        for a_name in param_names['A']:
            try:
                fit.parameters = {
                    param_names['y0'][0]: y0_guess,
                    a_name: A_guess,
                    param_names['t1'][0]: 5.0
                }
                fit.fit()
                result = fit.result()
                success = True
                break
            except:
                continue

        if not success:
            warnings.warn(f"Не удалось выполнить аппроксимацию для файла {filename}")
            continue


        # Получаем результаты с проверкой наличия параметров
        def get_param(result, names, default=None):
            for name in names:
                if name in result:
                    return result[name]
            return default


        t1 = get_param(result, param_names['t1'])
        t1_error = get_param(result, ['e_' + n for n in param_names['t1']])
        A = get_param(result, param_names['A'])
        A_error = get_param(result, ['e_' + n for n in param_names['A']])
        y0 = get_param(result, param_names['y0'])
        y0_error = get_param(result, ['e_' + n for n in param_names['y0']])
        r_squared = result.get('r', 0) ** 2 if 'r' in result else None
        niter = result.get('niter', 0)

        # Сохраняем результаты
        results.append({
            'Filename': filename,
            't1': t1,
            't1_error': t1_error,
            'y0': y0,
            'y0_error': y0_error,
            'A': A,
            'A_error': A_error,
            'R_squared': r_squared,
            'Iterations': niter
        })

        print(f"Успешно обработан: {filename}")
        print(f"Результаты: t1={t1:.4f}±{t1_error:.4f}, A={A:.4f}±{A_error:.4f}, R²={r_squared:.4f}")

    except Exception as e:
        warnings.warn(f"Ошибка при обработке файла {filename}: {str(e)}")

# Сохраняем все результаты в CSV
if results:
    results_df = pd.DataFrame(results)
    # Упорядочиваем столбцы
    cols = ['Filename', 't1', 't1_error', 'A', 'A_error', 'y0', 'y0_error', 'R_squared', 'Iterations']
    results_df = results_df[cols]
    results_df.to_csv(output_file, index=False)
    print(f"\nРезультаты сохранены в {output_file}")
    print(results_df)
else:
    print("\nНе удалось обработать ни один файл")

op.exit()