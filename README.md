main. Task for checking sector_keys with weak coverage zones and HLD.

1. Выгружаем results с емкостного калькулятора beeplan
2. В файл Enhance_req_GU(res_capacity в python) на лист 2100_newsite выносим те GU, где требуется расширение и только шаги 2100-new_site
   Отдельно на соответствующие листы выносим шаги 2100 --> new site. 
   Которым требуется new_site - колонка DS enhancerequired=1, и так по уменьшению до 2100.
3. В EcellsList формируем связку eNodeB ID-Cell ID и получаем привязку eNodeB ID-Cell ID-sector_key
4. В файлах Weak_Cov:
   3.1 выбираем кол-во измерений чуть больше среднего или среднее, и % плохого покрытия > 5%;
   3.2 по eNodeB ID-Cell ID в EcellsList находим sector_key
   3.3 по sector_key в Enhance_req_GU (res_capacity в python) находим sector_key_enh
   3.4 оставляем только те, где sector_key_enh имеет значения.
5. Из файла Enhance_req_GU выбираем те строки, которые есть в п.3.4 и вставляем их в HLD_емкость на листы sites и entrances соответственно.
   Для листа sites файла HLD_емкость из файла results берем данные с листа merge result, объединяем в одну строку.