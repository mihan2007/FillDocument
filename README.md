Этот проект представляет собой решение для автоматизации заполнения тендерных документов, что существенно упрощает процесс, особенно когда требуется ввод информации о разных юридических лицах. Вам больше не придется копировать вручную копировать в одном документы и вставлять в другой документ данные такие об организации, такие как наименование, реквизиты, ИНН и пр. После интеграции в Word макросы можно будет добавить на панель быстрого доступа или же на сочетание клавиш. 

**Основные функции:**

1. **DefaultCompany.xml и CompanyList.xml:** 
   - **DefaultCompany.xml** - файл, в котором хранятся данные выбранной организации. Эти данные скопированы из **CompanyList.xml** и будут автоматически внесены в документы.
   - **CompanyList.xml** - это своего рода база данных, содержащая информацию о юридических лицах, с которыми вы работаете.

2. **LegalEntityProfile - Редактор Профилей:**
   - Этот редактор позволяет вам редактировать, удалять и добавлять организации в список.!
![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic1.png)
3. **Выбор организации:**
   - Вы можете выбрать нужное юридическое лицо из списка в верхней части экрана.
   ![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic2.png)
   -При нажатии на кнопку "Выбрать", данные организации из файла **CompanyList.xml** будут скопированы в **DefaultCompany.xml**, и название выбранной организации будет отображаться под списком. Перед заполнением форм убедитесь, что       выбрано именно то юр. лицо, которое вам нужно
![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic3.png)
4. **Создание новой записи:**
   - При нажатии кнопки "Создать", форма будет очищена, и вы сможете вводить данные. После завершения ввода данных, появится диалоговое окно, предлагающее создать запись.

5. **Редактирование данных:**
   - Нажав кнопку "Редактировать", текстовые поля станут доступными для редактирования, и изменения будут сохранены.

**Интеграция в Word:**

Для интеграции этой программы в Microsoft Word:

1. Скачайте файлы в папку на вашем диске (например, "C:\Tenders\").

2. Откройте Word и перейдите в "Файл", затем в "Настройки".

![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic5.png)

3. Выберите "Настройка ленты" и убедитесь, что у вас установлена галочка рядом с "Разработчик".

![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic6.png)

4. Закройте окно настроек. Теперь в лентах Word у вас должна появиться вкладка "Разработчик". Перейдите в нее и нажмите "Visual Basic".

![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic7.png)

5. В редакторе макросов выберите "Normal" и импортируйте файлы, выбрав их из папки, где вы сохранили файлы.

![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic8.png)

![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic9.png)

в итоге должно получится вот так 

![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic10.png)

6. Далее нажимаем на LoadCompanyProfile

![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic11.png)

В этих строках прописываем место расположение  файлов -  CompanyList.xml и  DefaultCompany.xml, 
Public Const MainXMLFilePath As String = "C:\Tenders\CompanyListInfo.xml"
Public Const ChoisenCompanyXMLFilePath As String = "C:\Tenders\DefaultCompany.xml"

если вы последовали моему совету и с копировали файлы в "C:\Tenders\ то Вам менять ничего не надо. 

7. Откройте модуль "ShowCompanyProfile" и убедитесь, что все работает. Вызовите макрос "RunProgramm".

![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic12.png)

Дале можете поставить курсор на макросы ниже и проверить работают они или нет. Далее опять идем в Настройки Word настройки ленты, и создаем отдельную вкладку с макросами, как показано в примере 

8. Создайте отдельную вкладку с макросами в настройках Word, чтобы управлять программой. Настройте команды "Макросы" и добавьте "RunProgramm" в созданную вкладку.

![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic13.png)

В разделе команды выбираем «Макросы», выбираем нужный RunProgramm  макрос и жмем кнопку добавить, макрос  должен добавится в вновь созданную группу. Название макроса получились слишком длинные и не влезет полностью в окно, поэтому необходимо ненадолго задержать курсор на против каждого макроса из списка и появится всплывающее меню  с названием макроса.  
При желании можете настроить выявление макросов на сочетания клавиш
Теперь появившейся вкладке будет отображаться выбранный макрос при нажатии на него появится редактор профилей.

![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic14.png)

9. Настройте панель быстрого доступа, добавьте макросы и настройте иконки по вашему усмотрению.

![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic15.png)

Теперь, при нажатии на иконки или использовании сочетаний клавиш, вы можете легко вставлять данные организаций в документы.

![pictures/pic1.png](https://github.com/mihan2007/FillDocument/blob/main/pictures/pic16.png)
