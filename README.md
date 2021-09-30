# HR-Helper
Программа-помощник  для HR.    
**Последнюю версию программы можно скачать здесь https://github.com/praudmur/HR-Helper/releases**

  

![изображение](https://user-images.githubusercontent.com/1259834/135495987-1e70c389-5d05-487f-aee8-85136929c2c6.png)

Disclaimer
---
Программа не требует подключения к интернету.<br/>
Программа не отправляет письма сама, а лишь формирует черновики в MS Outlook.<br/>
На данный момент программа работает только с MS Outlook.<br/>
Чем можно заменить программу:
- Mail Merge в MS Word, но он не позволяет автоматически вкладывать вложения в письма.
- Mail Merge с вложениями в MS Publisher, но он не даёт такой гибкости, как программа HR Helper.

**Инструкция**
---
**I.Добавление базы работников**
---
Для начала необходимо загрузить базу работников из Excel файла.
1) Пустой шаблон для базы можно выгрузить через меню «Файл->Выгрузить шаблон базы работников».

![изображение](https://user-images.githubusercontent.com/1259834/135496624-744e10b9-bc7b-4c09-8d1d-931a98442bed.png)

Файл Excel имеет 3 обязательных столбца – «ФИО», «Email_Work», «Email_Home». Остальные столбцы добавляются по желанию(примеры использования доп. Столбцов см. ниже) 

![изображение](https://user-images.githubusercontent.com/1259834/135496669-dea502f1-4074-4bf6-850c-c67a24fe89ef.png)

2)Заполним тестовую базу(все ФИО и адреса сформированы случайным образом)

![изображение](https://user-images.githubusercontent.com/1259834/135496696-4fe3d788-0da8-48f5-965a-adcea633c244.png)

3)Нажимаем кнопку «Добавить базу» и выбираем файл с базой.

![изображение](https://user-images.githubusercontent.com/1259834/135496724-fd9c18e3-217e-4371-b0bf-8ab1f2c64016.png)  
4)Готово! В списке сотрудников должны появиться Ваши работники. Можно отметить галочкой тех, кому нужно отправить письмо и нажать кнопку «Начать рассылку» для формирования писем.

![изображение](https://user-images.githubusercontent.com/1259834/135496749-9ea4fdcc-635e-476f-8b20-27ced7beba0f.png)

**II.Переменные в теле письма.**
---
HR Helper позволяет использовать в теле письма переменные по типу Mail Merge в MS Word. Например, мы можем добавить переменную «имя» и для каждого работника письмо будет содержать именно его имя.  
1)Добавляем столбец «Имя» в наш Excel файл.  
![изображение](https://user-images.githubusercontent.com/1259834/135496776-8ca4eea7-b1ea-438d-9d14-cd5a577a4bb7.png)  
2) Если база уже загружена в программу, нужно заново нажать «Добавить базу»  
3)В теле письма добавляем названием переменной в <> скобках, например:  
![изображение](https://user-images.githubusercontent.com/1259834/135496789-39972363-9211-4258-a498-71a19013d93e.png)  
4)В письме получаем:  
![изображение](https://user-images.githubusercontent.com/1259834/135496814-a4b4bd20-d716-4da7-b18c-d271bcc4bfb2.png)  

Переменных можно добавлять неограниченное количество, также они могут содержать в себе неограниченный объём текста. При желании тело письма может состоять из переменных.


**III.Вложения в письма**
---
  
HR Helper позволяет очень гибко работать с вложениями в письмах.  
  
**III.1 – Простые вложения в письма**
---
Эти вложения будут добавляться во все письма при формировании.  
1)Ставим галочку в поле «Добавить выбранные файлы».  
![изображение](https://user-images.githubusercontent.com/1259834/135496925-44027b61-7d21-4f64-813e-61791290af30.png)  
2)Нажимаем кнопку «Добавить вложения» и выбираем нужный файл  
![изображение](https://user-images.githubusercontent.com/1259834/135496947-9bfe4228-d1fe-4c25-8f24-3520599bd153.png)  
3)Он появится в списке вложений  
![изображение](https://user-images.githubusercontent.com/1259834/135496972-abb141bd-19d5-4ff5-8165-942a8c22ecab.png)  
4) При формировании писем этот файл будет добавляться всем работникам.  


**III.2 – Вложения в письма по имени работника**
---
В названии файлов программа будет искать имя работника и добавлять файл ему в письмо. Полезно, если у Вас есть папка, которая содержит файлы для большого количества работников.
1)Нам нужно отправить работникам «Иванов Иван Иванович» и «Королева Александра Матвеевна» письма с индивидуальным вложением.  
2) Ставим галочку в опции «Искать вложения».  
![изображение](https://user-images.githubusercontent.com/1259834/135497017-756633d0-d4ff-48a4-bdbf-cf637be96a62.png)  
3) Нажимаем кнопку «Выбрать папку вложений» и выбираем папку, где находятся наши файлы  
![изображение](https://user-images.githubusercontent.com/1259834/135497033-93ec5738-f083-4863-88b5-d21c3910ceeb.png)  
![изображение](https://user-images.githubusercontent.com/1259834/135497042-4ac6a128-2dd5-4136-ac3a-cecc78ab0899.png)

4) Нажимаем кнопку «Начать рассылку»  

6)В письме для каждого работника будет прикреплено своё вложение. Поиск происходит по ФИО  
![изображение](https://user-images.githubusercontent.com/1259834/135497082-0e55722e-9ac6-47ba-8879-534dbb1e807d.png)  
![изображение](https://user-images.githubusercontent.com/1259834/135497096-b2c1c471-eae1-4657-b7e4-7e2f05b95117.png)  

**III.3 – Вложения в письма из тела письма**
---
Программа даёт возможность добавлять вложения в письма в самом теле письма.  
1) Ставим галочку в опции «Искать вложения».  
![изображение](https://user-images.githubusercontent.com/1259834/135497144-0bdd4abe-fb3a-4c7a-98bf-e0c51cd96fce.png)  
2) Нажимаем кнопку «Выбрать папку вложений» и выбираем папку, где находятся наши файлы  
![изображение](https://user-images.githubusercontent.com/1259834/135497153-76600b84-d10d-4c79-a69b-2f4eba73ae3c.png)  
![изображение](https://user-images.githubusercontent.com/1259834/135497162-8f2d8eb9-ee5a-4ac8-ae31-fe14f59a77fb.png)  
3)В тело письма добавляем название файла в скобках <вложить></вложить>, например <вложить>Разослать_Файл.txt</вложить>,  
![изображение](https://user-images.githubusercontent.com/1259834/135497177-fe63d700-ae47-48d8-b3af-3630ba50c762.png)  
4)В письме будет приложен указанный файл(если он существует в выбранной папке вложений) и текст <вложить>Разослать_Файл.txt</вложить> будет автоматически удалён.  
![изображение](https://user-images.githubusercontent.com/1259834/135497198-ea917c67-c7a4-4ec0-829e-bcf3a2e7b1ec.png)  
5)Для указания файла можно использовать переменные (см. раздел II инструкции).  
Например, добавляем столбец в файл Excel  

![изображение](https://user-images.githubusercontent.com/1259834/135497211-b57580f7-8e4c-4549-8cb5-e0308890a3f2.png)  
В теле письма добавляем <вложеннный файл>  

![изображение](https://user-images.githubusercontent.com/1259834/135497227-e19a162c-6d48-4183-a6dc-fb46bb36175b.png)  

Для каждого работника в письме добавляется вложение «Разослать_Файл1.txt» или «Разослать_Файл2.txt» в зависимости от его значения в столбце <вложеннный файл>.  




Как скомпилировать программу самостоятельно:
---
- Добавить mso.dll и lib\msoutl.olb из папки Microsoft Office в **папку проекта\lib\**
- Изменить путь к библиотекам boost в **C/C++->General->Additional Include Durectories**
- Распаковать wxWidgets в папку **папку проекта\lib\**
- Построить библиотеку OpenXLSX, OpenXLSX-static.lib положить в **папку проекта\lib\**


Благодарности
---
troldal за OpenXLSX  
Разработчикам wxWidgets за wxWidgets  
brofield за simpleini  

---


BSD 3-Clause License

Copyright (c) 2021, praudmur
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its
   contributors may be used to endorse or promote products derived from
   this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

