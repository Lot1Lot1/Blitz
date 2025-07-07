import originpro as op
import pandas as pd

# 1. Загрузка данных
data = pd.read_csv("ZE15_250_Ag_dis.csv")
ws = op.new_sheet('w')
ws.from_df(data)

# 2. Настройка аппроксимации
fit = op.NLFit('ExpDecay1')
fit.set_data(ws, 0, 1)  # X=0, Y=1

# 3. Фиксация параметров (ключевая часть!)
fit.fix_param('y0', 0)       # Жестко фиксируем y0=0
fit.set_param('A', 1.0)      # Начальное значение амплитуды
fit.set_param('t1', 5.0)     # Начальное время затухания

# 5. Запуск аппроксимации
fit.fit()

# 6. Вывод результатов
result = fit.result()
print(f"t1 = {result['t1']:.4f} ± {result['e_t1']:.4f}")
