   M S B A M L   `   `   `  ���� I  EDataServiceApp, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null/clr-namespace:DataServiceAppDataServiceApp      DataServiceApp.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     .localclr-namespace:DataServiceApp     -     � Title$  Главное окно��$	��450��$��390.909�� � 
ResizeMode$ CanMinimize=� � WindowStartupLocation$ CenterScreen=�.��� .+��� .@��� -    ��Name$ stateRecyclerTab������ $��Управление����� ���	#FF91C4DE.+��� .+��� $��Инсталляция��$��10q�.���� .+��� -    ��Name$ 
btnInstall��$��Установить���� -   $ 	btnDelete��$��Удалить���� $��Запуск��$��10q�.���� .+��� -   $ btnStart��$��Запустить���� -   $ 
btnReStart��$ ��Перезапустить���� -   $ btnPause��$ ��Приостановить���� -    ��Name$ txtCurrentState��$6��0Текущее состояние службы: ��$��Center=�$��14(�$��0,20,0,0q� �� TextAlignment$ Center=��� -	   $ stateSettingsTab������ $'��!Параметры запуска��.��� ���	#FF91C4DE � ColumnDefinitions �� $
��176* ��� $
��203* � � RowDefinitions �� $	��33* ��� $	��58* ��� $	��47* ��� $
��227* ��� $	��28* �.+��� $)��#*Путь к базе данных:��$��2��$��15(�$��4��$��10,30,10,191q�.��� .+��� -
   	 ��Name$	 tbPath��$��10,10,10,72q��� -   $ btnApply��$	��100��$��Right=�$��Применить��$
��Auto��$��
0,60,10,10q�
 J� RenderTransformOrigin$
 
0.46,0.189&��� -   $ 	btnSelect��$	��100��$
��Left=�$��Выбор..��$
��Auto��$��
10,60,0,10q��� $:��4*Указывать папку, а не .bd файл��$��15(�$��33��$��2��$��10,212,5,10q�$��3��$��2���� $=��7Частота обновления (в минутах)��$��15(�$��2��$��10,73,96,33q�$��3��.��� .+��� -   $	 tbFreq��$��10,10,90,50q��� -   $ btnApplyFreq��$	��100��$
��Left=�$��Применить��$��	10,49,0,0q�$��38��$	��Top=��� -   $ 
tiSettings������ $��Настройки��.���� .+��� -    ��Name$ cbAppWindows��$4��.Включать при запуске Windows��$��
10,10,0,10q�$��14(��� -   $ cbServiceStart��$<��6Автоматический запуск службы��$��
10,10,0,10q�$��14(��� -   $ cbTrayMinimize��$(��"Сворачивать в трей��$��
10,10,0,10q�$��14(�