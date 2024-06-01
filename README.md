# ПРАКТИЧНА РОБОТА №3
## ВИКОРИСТАННЯ COM-ТЕХНОЛОГІЇ. ЕКСПОРТ ДАНИХ У ДОКУМЕНТИ WORD ТА EXCEL
**Мета роботи:** Вивчення та практичне відпрацювання механізмів використання компонентів на основі СОМ -технології та її застосування в .NET. Вивчення особливостей використання додатків та програмного формування документів MS Office.
## Результат виконання роботи
**Варіаційна частина:** Приватна клініка, було взято список записів користувачів на прийом.  

Створено 3 проєкти в одному рішенні:   
![image](https://github.com/JuliaSylenok/Lab_3_Sylenok/assets/149322465/61c0a768-da5a-483c-b860-d1c76270ff44)


Результат створеного Excel файлу:  
![image](https://github.com/JuliaSylenok/Lab_3_Sylenok/assets/149322465/320f93e1-9ef9-4a09-b101-be15984a0e27)
  
Результат створеного Word файлу:  
![image](https://github.com/JuliaSylenok/Lab_3_Sylenok/assets/149322465/b560560c-d3ee-47a2-b882-39de559e9dcc)

| № | Вимоги до роботи  | Бали | Що виконано| 
---| ------------- | ------------- |---|
1| Експорт текстових даних у документ MS Word та коректне завершення роботи COM -об'єкта  | 2  | +
2| Експорт текстових даних до таблиці MS Excel та коректне завершення роботи COM -об'єкта  | 2  | +
3| Реалізація завдань експорту (п.1 і п.2) як окремих збірок (dll assembly ), що динамічно завантажуються.  | 1  | +
4| Додатково до п.1 програмне формування таблиць для даних MS Word, що експортуються. | 2  | +
5| Пізніше зв'язування. Налаштування активного модуля експорту користувачем під час роботи програмного забезпечення. | 1  | +

## Висновок  
В результаті виконання роботи було створено два класи Library - ExcelExporter та WordExporter, які формують звіти у форматах MS Word та MS Excel з використанням COM-технологіхї для взіємодії з відповідними програмами Microsoft Office.
