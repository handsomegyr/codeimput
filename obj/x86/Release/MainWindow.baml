   M S B A M L   `   `   `  ���� M  ISystem, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089M3clr-namespace:System.ComponentModel;assembly=SystemSystem.ComponentModel   
CodeImport"   ExcelApplication1.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     Bcm3clr-namespace:System.ComponentModel;assembly=System     -     � Title$  卡券导入工具��$	��768��$
��1024�� � WindowStartupLocation$ CenterScreen=� � WindowStyle$ ThreeDBorderWindow=� � 	Resources ,   &System.ComponentModel.BackgroundWorker   backgroundWorker   WorkerReportsProgress!   WorkerSupportsCancellation��%"   &	          -    .  . .���  � RowDefinitions �� $��50 �$��50���� $��30 �$��30���� $	��200���� $��50 �$��50�� � ColumnDefinitions �� .+�� -   $��grid1��$��Stretch=�$��Stretch=� � ShowGridLines .  $��0�� �� $	��100��$	��16* ��� $	��45* ��� $	��100��$
��auto �.+��� -   $��button2��$��选择��$��22��$
��Left=�$��Center=�$��	7,12,0,16q�$��42��$��2���� -   $��
textBlock1��$��Stretch=�$��请选择一个目录:��$��Center=�	 �� TextAlignment$	 Justify=�
 �� Padding$
 10,0,0,0q�$��0��$��2��$
��0,17q��� -    ��Name$ button2_refresh��$��刷新��$��22��$��Center=�$��Center=�$��
54,12,5,16q�$��42��$��2���� $��1�� �� Orientation$ 
Horizontal=�.+��� -   $��
textBlock2��$��Stretch=�$�� ��$��Center=�$	 Justify=�$
 10,0,0,0q��� -   $��
textBlock3��$��Stretch=�$�� ��$��Center=�$	 Justify=�$
 30,0,0,0q��� -	   $��listBox1��$��Stretch=�$��
12,14,12,0q�$��Stretch=�$��Stretch=�$��2��$	��180��$	��480��$��Stretch=����� .R�� $��3q� �� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto �.+��� $��0��$��3q�$��Center=� d� 	IsChecked ��  �� Path$ 
IsSelected� �� Mode$ TwoWay=��� $��1��$
 5q�$��Center=����� $
 Name��� $��2��$
 6q�$��Center=����� $
 Path��� $��3��$
 3q�$��Center=����� $ Status�� -
   $��grid2��$��Stretch=�$��Stretch=� .  $��3�� �� $	��100��$��* ��� $	��100��$
��auto ��� $	��100��$
��auto �.+��� -   $��	cmdCancel��$��取消��$��23��$��Right=�$��0,5,8,5q�$��80��$��2��$��Stretch=��� -   $��button1��$��合并&导出��$��23��$��Right=�$��0,5,8,5q�$��80��$��1��$��Stretch=�� -   $��progressBar1��$��22��$��Stretch=�$��11,15,10,13q�$��Center=�