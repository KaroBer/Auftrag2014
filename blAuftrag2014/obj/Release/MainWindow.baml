   M S B A M L   `   `   `  ���� O  Kmscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e0891&clr-namespace:System;assembly=mscorlibSystem   blAuftrag2014   blAuftrag2014.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     9System&clr-namespace:System;assembly=mscorlib       � Title$  blAuftrag2014��$	��620��$	��840�� � 	Resources    ListenZeilen �� Property �� Value �� 
TargetType �� BasedOn �� Triggers P� Property P� Value��%3  &	        '�� �     7   �� .���  � ��$ 1q���  � ��$ Consolas)���  � ��$ 14(���  � ��$ DemiBold"���  � ��$ Bottom=���  � ��
 ��  ���  � ��
 ������� " ��8    P�  �@J�IsMouseOver . . ���  � ��
 �������  � ��
 �   �.��� .+��� $
��Auto=�$��0q�.���� $��0q����� 	 �� ImageSource$	 zwergenhaeuser.jpg��.+��� -   
 ��Name$
 TabControl1��$��0,0,0,0q�
������ $��2��.@��� -    ��Name$ ti_Aufträge��$��	Aufträge��.���� .+��� .+���  �� Orientation$ 
Horizontal=�$��
0 -28 10 0q�$��Right=�.+��� -   $��	lbl_heute��$��Freitag, 20.11.2015��
������ 
�������$��Sketchflow Print)�$��18(�$
��Bold"��� $ 
Horizontal=�.+��� -    ��Name$ lbl_dnp��$��	20 12 0 4q�$	��dnp��$��80��
�������
���  ��$��Cooper Std Black)�$��18(�$��Center=�
���  ��$��6q��� -   $ lbl_wum��$��	0 12 10 4q�$	��wum��$��80��
������ 
��� � �$��Cooper Std Black)�$��18(�$��Center=�
���  ��$��6q��� -   $ lbl_HeftNrminus1��$��20 10 10 10q�$
��1404��
������ 
��� � �$��Cooper Std Black)�$��18(�$��90��$
��Left=�$��Center=�$��Bottom=�
���  ��$��3q� K� ZIndex$ 1���� -   $ lbl_HeftNrAktuell��$��
-30 6 10 6q�$
��1405��
�������
���  ��$��Cooper Std Black)�$��24(�$	��110��$
��Left=�$��Center=�
���  ��$��6q�$ 2���� -   $ lbl_HeftNrplus1��$��-30 10 10 10q�$
��1406��
������ 
��� d �$��Cooper Std Black)�$��18(�$��90��$��Bottom=�$
��Left=�$��Center=�
���  ��$��3q�$ 1���� -	    ��Name$ tb_sucheAuftraege��$��138 -10 0 0q�$	��130��$��26��
����  �$��2q��� -
    ��Name$ bt_suche��$��	0 -10 0 0q�$��60��$��24��$��Suche��$��14(��� $ 
Horizontal=�$��0 -3 0 0q�.+��� -   $ lbl_Zahl��$��30 -5 0 -10q�$��17��$	��Top=�$
��left=�$��Center=�$��60��
������ 
�������$��Sketchflow Print)�$��40(�$
��Bold"��� -   $ lbl_com��$��10 0 0 0q�$	��com��$	��Top=�$
��left=�$��Center=�$��80��
������ 
��� � �$��Cooper Std Black)�$��18(�
���  ��$��6q��� -   $ bt_RG_vorbereiten��$��Rechnung vorbereiten��$	��140��$��24��$��	55 6 0  0q��� -   $ bt_NeuerAuftrag��$��	282 6 0 0q�$
��left=�$	��140��$��24��$��Neuer Auftrag��$��14(��� -    ��Name$ stp_NeuerAuftrag��$y�	Collapsed=�
����  �$��10q�$��	194 0 0 0q�$	��565��$ 2��._��� 
�������.+��� $ 
Horizontal=�$��10 0 0 0q�.+��� $
��Heft��$��90���� $��Ausgabe��$��96���� $
��Text��$	��186���� -   $��bt_typ��$	��Typ��$��40��$��20��$��0 0 60 0q��� $��je��$��50���� $ 
Horizontal=�$��10 -4 10 10q�.+��� -   $ tb_heft��$	��dnp��$��80��$��14(�$��0 0 0 0q��� -   $ 
tb_ausgabe��$
��1405��$��80��$��14(�$��8 0 10 0q��� -   $ tb_text��$��
Codekicker��$	��170��$��14(�$��8 0 10 0q��� -   $ tb_typ��$	��std��$��80��$��14(�$��8 0 10 0q��� -   $ tb_je��$��35��$��50��$��14(�$��8 0 0 0q��� -   $ bt_NeuenAuftragSichern��$��10 -5 10 10q�$��Auftrag sichern��$	��400��$��24��$��center=�$��14(��� -    ��Name$ stp_Listbox_etc��$ 
Horizontal=�
������ $��0 4 0 10q�$ 1��.+��� $��Bottom=�.+��� -    ��Name$ lb_auftragsliste��
������ $��2 0 0 0q�$	��232��$	��276��$
��left=�$	��Top=�$��0q� � ContextMenu �� .@�m� -   $��
Umbenennen��m� -   $��Löschen��m� -   $#��Ins nächste Heft verschieben��.@��� $��dnp1405Codekicker���� $��	0 -10 0 0q�$	��top=�.+��� -   $ lbl_AuftragsString��$��0 0 0 0q�$��16(�$	��562��$��35��$@��:dnp1405Codekicker -- 1,5 -- std -- je: 35 -- 52,5 -- 2 AEs��
����  �
�������$��
ExtraBlack"�$��Center=�
����  �$��2q��� -   $ Zeiterfassung��$y�Visible=�$��0q�.+��� .+��� -   $ lb_AEzeilenZeit��
�������$	��562��$	��226��$��0q�$
��Left=�$��Consolas)�.@���AEZeilenliste���� 
�������$ 
Horizontal=�$��0 0 0 0q�.+��� -   $ lbl_idxZeit��$�� ��$��26��$��1q��� $
��Text��$	��150��$��4 0 0 0q��� $��Datum��$��70��$��1q��� $	��von��$��40��$��1q�$��Center=��� $	��bis��$��40��$��1q�$��Center=��� $��Dauer��$��50��$��14 1 1 1q��� 
�������$ 
Horizontal=�$��0 0 0 0q�.+��� -    $ bt_neueEingabezeileZeit��$��20��$��C��$��	10 -6 1 6q��� -!   $ 	ez_tbtext��$	��150��$��1 -5 1 7q� � ToolTip$  Doppelklick = Standardtext��{�.  �� -"   $ 
ez_tbdatum��$��70��$��1 -5 1 7q�$��Center=�$ Doppelklick = heute��{�.  �� -#   $ ez_tbvon��$��40��$��1 -5 1 7q�$��Center=�$! Doppelklick = jetzt (09:30)���� -$   $ ez_tbbis��$��40��$��1 -5 1 7q�$��Center=�$! Doppelklick = jetzt (10:00)���� -%   $ ez_lbldauer��$��70��$��	-4 -7 1 7q�$��Center=��� -&   $ 
bt_zeit_OK��$
�� OK ��$��30��$y�Visible=�$��6 -4 0 6q��� -'   $ ez_lblmodus��$��
10 -12 1 1q�$
��Left=�$��18(�
��� d �$
��Bold"�$	��Neu���� -(   $ Stückerfassung��$y�	Collapsed=�.+��� .+��� -)   $ lb_AEzeilenST��
�������$	��560��$	��224��$��0q�$
��Left=�$��Consolas)�.@���FBListe AEZeilen Stück, der Startstring so ausgeweitet, dass er ...���� 
�������$ 
Horizontal=�$��0 0 0 0q�.+��� -*   $ lbl_idxStueck��$�� ��$��26��$��0q��� $
��Text��$	��200��$��4 -2 1 1q��� $��Datum��$��64��$��4 -2 1 1q��� $
��Wert��$��40��$��4 -2 1 1q�$��Center=��� 
�������$ 
Horizontal=�$��0 -2 0 0q�.+��� -+   $ bt_neueEingabezeileStueck��$��20��$��C��$��	10 -5 1 6q��� -,   $ 	st_tbtext��$	��200��$��2 -5 1 6q�{�.  $ ToolTip���� --   $ 
st_tbdatum��$��70��$��2 -5 1 6q�$��Center=�$ Doppelklick = heute��{�.  �� -.   $ 	st_tbwert��$��50��$��2 -5 1 6q�$��Center=��� -/   $ bt_stueck_OK��$
�� OK ��$y�Visible=�$��4 -5 1 6q��� -0   $ st_lblmodus��$��
10 -12 1 6q�$��30��$
��Left=�$	��Top=�
��� d �$
��Bold"�$��18(�$	��Neu���� $ 
Horizontal=�$��4q�$	��790��$
��left=�
�������$ 2��.+��� $��Heute��$��14(�$��	24 0 -6 4q�$��Center=��� -1   $ lbl_tag��$	��122��
��� � �$��14(�$
��Bold"�$��60��$��-8 0 0 4q�$��Center=��� -2   $ lbl_tagSchnitt��$��(-19 | 141)��$��14(�$��80��$��
-18 0 12 4q�$��Center=��� $��Woche��$��14(�$��2 0 -6 4q�$��Center=��� -3   $ lbl_kw��$	��410��$��14(�$
��Bold"�$��60��$��	-10 0 0 4q�$��Center=��� -4   $ lbl_kwSchnitt��$��(-491 | 901)��$��14(�$��90��$��
-14 0 12 4q�$��Center=��� $��Monat��$��14(�$��2 0 -6 4q�$��Center=��� -5   $ 	lbl_monat��$
��2120��$��14(�$
��Bold"�$��60��$��-6 0 0 4q�$��Center=��� -6   $ lbl_monatSchnitt��$��(-1003 | 3123)��$��14(�$	��100��$��
-10 0 12 4q�$��Center=��� $
��Jahr��$��14(�$��2 0 -6 4q�$��Center=��� -7   $ lbl_jahr��$
��7580��$��14(�$
��Bold"�$��60��$��0 0 2 4q�$��Center=��� -8   $ lbl_jahrSchnitt��$��(-28877 | 36457)��$��14(�$	��115��$��	-10 0 0 4q�$��Center=��� $ 
Horizontal=�$��4 -8 0 0q�$��26��
�������$ 3��$
��Left=�$	��790��.+��� -9   $ lbl_Vortage��$��212 | 180 | 90��$��12(�$	��192��$��4,-4,0,2q�$��Center=��� -:   $ lbl_Vorwochen��$��1352 | 680 | 490��$��12(�$	��180��$��-25 -4 20 2q�$��Center=��� -;   $ lbl_Vormonate��$��2352 | 3680 | 2490��$��12(�$	��180��$��-10 -4 17 2q�$��Center=��� -<   $ lbl_Vorjahre��$��36512 | 43680 | 32490��$��12(�$	��180��$��0 -4 0 2q�$��Center=��� -=   $ ti_Laufzettel��$��
Laufzettel��.���� $��
56 2 240 0q�.+��  � ColumnDefinitions �� $��80 ���  � RowDefinitions �� �� �� �� �� �� �� �� �� �� �� �� �� .+��� $��1q�
�������$
��Bold"�$��0��$��0��$��Ausgabe��$��12(��� $��1q�
�������$
��Bold"�$��1��$��0��$��Rubrik��$��12(��� $��1q�
�������$
��Bold"�$��2��$��0��$��Thema��$��12(��� $��1q�
�������$
��Bold"�$��3��$��0��$��Seite��$��12(��� $��1q�
�������$
��Bold"�$��4��$��0��$��	Redakteur��$��12(��� $��1q�
�������$
��Bold"�$��5��$��0��$��Datum��$��12(��� $��1q�
�������$
��Bold"�$��7��$��0��$��Anmerk1��$��12(��� $��1q�
�������$
��Bold"�$��8��$��0��$��Anmerk2��$��12(��� $��1q�
�������$
��Bold"�$��9��$��0��$��Anmerk3��$��12(��� $��1q�
�������$
��Bold"�$��10��$��0��$��Anmerk4��$��12(��� $��0��$��1��$ 
Horizontal=�.+��� ->   $��	tbAusgabe��$	��100��$��16(�$��-2015���� -?   $��plus1��$��+1��$��30���� -@   $��	lfz_clear��$��Clear��$��50��$��	250 0 0 0q��� -A   $��tbRubrik��$��1��$��1��$��16(�$�� ���� -B   $��tbThema��$��2��$��1��$��16(�$�� ���� -C   $��tbSeite��$��3��$��1��$��16(�$�� ���� -D   $��tbRedakteur��$��4��$��1��$��16(�$��la���� -E   $��tbDatum��$��5��$��1��$��16(�$�� ���� -F   $
��tbA1��$��7��$��1��$��16(�$�� ���� -G   $
��tbA2��$��8��$��1��$��16(�$�� ���� -H   $
��tbA3��$��9��$��1��$��16(�$�� ���� -I   $
��tbA4��$��10��$��1��$��16(�$�� ���� -J   $��tb_zielpfad��$��11��$��1��$��2��$��16(�$��Zielpfad���� -K   $��tb_zieldatei��$��12��$��1��$��16(�$��	0 0 0 120q�$��PDFName���� -L   $��bt_PDFschreiben��$��	0 0 0 120q�$��11��$��2��$��0��$��Go��$��20(�$
��Bold"��� -M   $ ti_Rechnung��$��Rechnung��.���� $	��790������ $	 zwergenhaeuser.jpg��.+��� $ 
Horizontal=�$
��left=�.+��� -N   $ 
bt_details��$��Details��$��12(�$��40 2 0 2q�$��20��$��80���� -O   $ bt_RGfürListe��$�� ... ��$��12(�$��
40 2 380 2q�$��20��$	��140���� -P   $ 
bt_kleiner��$	�� - ��$��14(�$��Consolas)�$��2q��� -Q   $ lbl_fontsize��$��14��$��2q��� -R   $ bt_groesser��$	�� + ��$��14(�$��Consolas)�$��2 2 40 2q� J� RenderTransformOrigin$ 	0.507,0.5&��� -S   $ 
tb_RGtexte�� �� AcceptsReturn . $	��760��$	��410��$
��left=�$��12 0 0 0q�" �� VerticalScrollBarVisibility$
 Auto=�$��14(�$��Consolas)��� -T   $ bt_Clipboard��$��	20 4 30 0q�$��14(�$��24��$	��560��$*��$Inhalt in Zwischenablage übernehmen��$��Center=��� -U   $ ti_Statistik��$��	Statistik��.���� $	��780������ .���  � Color$ Orange�� � Offset$	 0.0Q�� $ Cornsilk��$
 0.45Q�� $ Orange��$	 1.0Q�.+��� $ 
Horizontal=�$
��Left=�$	��800��.+��� -V   $ 
bt_Gestern��$��Gestern��$��12 4 4 4q�$��60���� -W   $ bt_statistik_Tag��$��Heute��$��4q�$��60���� -X   $ bt_Vorwoche��$��Vorwoche��$��4 4 4 4q�$��60���� -Y   $ bt_statistik_KW��$��Diese Woche��$��4 4 20 4q�$��80���� -Z   $ tb_wievielMonate��$��1��$��4q�$��20���� -[   $ bt_statistik_Monat��$��Monat(e)��$��2 4 2 4q�$��80���� -\    ��Name$ cbx_kurz��$��8 7 20 4q�$
��kurz�� d� 	IsChecked$
 trueY��� -]   $ tb_Matrix_von��$
��1410��$��4 4 2 4q�$��36���� -^   $ tb_Matrix_bis��$
��1502��$��2 4 2 4q�$��36���� -_   $ 	bt_Matrix��$��Matrix��$��2 4 26 4q�$��60���� -`   $ bt_kl��$	�� - ��$��14(�$��Consolas)�$��2 4 0 4q��� -a   $ lbl_fontsze��$��16��$��2q��� -b   $ bt_gr��$	�� + ��$��14(�$��Consolas)�$��2 4 55 4q�$ 	0.507,0.5&��� $ 
Horizontal=�$	��780��.+��� -c   $ tb_Statistik��$��12 2 2 2q� . $	��760��$	��440��$
 Auto=�$��16(�$��Consolas)��� -d   $ 	ti_Listen��$��Listen��.���� $ 
Horizontal=�$	��Top=�$	��500��$	��790������ $	 zwergenhaeuser.jpg��.+��� $
��left=�$��4 12 2 2q�.+��� -e   $ 	tb_Liste1��$
��Left=� . $	��350��$	��244��$
 Auto=�$��14(�$��Consolas)��� -f   $ 	tb_Liste2��$
��Left=�$��0 4 0 0q� . $	��350��$	��230��$
 Auto=�$��14(�$��Consolas)��� .+��� -g   $ 	tb_Liste4��$��2 12 0 0q�$
��Left=� . $	��424��$	��120��$
 Auto=�$��14(�$��Consolas)��� $ 
Horizontal=�.+��� -h   $ 	tb_Liste3��$��2 4 0 0q� . $	��180��$	��354��$	��Top=�$
��Left=�$
 Auto=�$��14(�$��Consolas)��� -i   $ 	tb_Liste6��$��4 4 0 0q� . $	��240��$	��354��$	��Top=�$
��Left=�$
 Auto=�$��14(�$��Consolas)��� -j   $ ti_Hochrechnung��$��
Überblick��.���� ���� $	 zwergenhaeuser.jpg��.+��� $ 
Horizontal=�.+��� -k   $��stp_tag��$
��Left=�$��60 6 2 2q�.+��� -l   $��	lbl_x_Tag��$	��Tag��$��Segoe)�$��16(�$
��Bold"�
���   �
�������$��Center=�$��Center=�$��2q�� $	��150�� �� ��  �� �� .+��� $��0��$��0��$ 
Horizontal=�
�������$��2q�.+��� $	��✔��$��Segoe UI Symbol)�$��16(�$
��Bold"�
��� � ��� -m   $��lbl_x_TagIst��$	��184��$��Segoe)�$��16(�$
��Bold"�
��� � �$��Bottom=��� $��1��$��0��$ 
Horizontal=�
�������$��2q�.+��� $��Ø��$��Segoe UI Symbol)�$��16(�$
��Bold"�
����  ��� -n   $��lbl_x_TagSchnitt��$	��145��$��Segoe)�$��16(�$
��Bold"�
����  �$��Bottom=��� $��0��$��1��$ 
Horizontal=�
�������$��2q�.+��� -o   $��lbl_x_TagAbw1��$��39��$��Segoe)�$��16(�$
��Bold"�
��� � �$��Bottom=��� $��1��$��1��$ 
Horizontal=�
�������$��2q�.+��� $�� ��$��Segoe)�$��20(�$
��Bold"�
���  ��$��Bottom=��� -p   $��	stp_woche��$
��Left=�$��20 6 2 2q�.+��� -q   $��lbl_x_Woche��$��Woche��$��Segoe)�$��16(�$
��Bold"�
���   �
�������$��Center=�$��Center=�$��2q�� $	��150�� �� �� ��  �� �� .+��� $��0��$��0��$ 
Horizontal=�
�������$��2q�.+��� $	��✔��$��Segoe UI Symbol)�$��16(�$
��Bold"�
��� � ��� -r   $��lbl_x_WocheIst��$	��386��$��Segoe)�$��16(�$
��Bold"�
��� � �$��Bottom=��� $��1��$��0��$ 
Horizontal=�
�������$��2q�.+��� $��Ø��$��Segoe UI Symbol)�$��16(�$
��Bold"�
����  ��� -s   $��lbl_x_WocheSchnitt��$	��725��$��Segoe)�$��16(�$
��Bold"�
����  �$��Bottom=��� $��2��$��0��$ 
Horizontal=�
�������$��2q�.+��� $	��⇗��$��Segoe UI Symbol)�$��16(�$
��Bold"�
�������� -t   $��lbl_x_WocheHR��$	��890��$��Segoe)�$��16(�$
��Bold"�
������$��Bottom=��� $��0��$��1��$ 
Horizontal=�
�������$��2q�.+��� -u   $��lbl_x_WocheAbw1��$
��-369��$��Segoe)�$��16(�$
��Bold"�
���  ��$��Bottom=��� $��1��$��1��$ 
Horizontal=�
�������$��2q�.+��� $�� ��$��Segoe)�$��20(�$
��Bold"�
���  ��$��Bottom=��� $��2��$��1��$ 
Horizontal=�
�������$��2q�.+��� -v   $��lbl_x_WocheAbw2��$
��+165��$��Segoe)�$��16(�$
��Bold"�
��� � �$��Bottom=��� -w   $��	stp_monat��$
��Left=�$��20 6 2 2q�.+��� -x   $ lbl_x_Monat��$��Monat��$��Segoe)�$��16(�$
��Bold"�
���   �
�������$��Center=�$��Center=�$��2q�� $	��150�� �� �� ��  �� �� .+��� $��0��$��0��$ 
Horizontal=�
�������$��2q�.+��� $	��✔��$��Segoe UI Symbol)�$��16(�$
��Bold"�
��� � ��� -y   $ lbl_x_MonatIst��$
��2687��$��Segoe)�$��16(�$
��Bold"�
��� � �$��Bottom=��� $��1��$��0��$ 
Horizontal=�
�������$��2q�.+��� $��Ø��$��Segoe UI Symbol)�$��16(�$
��Bold"�
����  ��� -z   $ lbl_x_MonatSchnitt��$
��2506��$��Segoe)�$��16(�$
��Bold"�
����  �$��Bottom=��� $��2��$��0��$ 
Horizontal=�
�������$��2q�.+��� $	��⇗��$��Segoe UI Symbol)�$��16(�$
��Bold"�
�������� -{   $ lbl_x_MonatHR��$
��2780��$��Segoe)�$��16(�$
��Bold"�
������$��Bottom=��� $��0��$��1��$ 
Horizontal=�
�������$��2q�.+��� -|   $ lbl_x_MonatAbw1��$
��+182��$��Segoe)�$��16(�$
��Bold"�
��� � �$��Bottom=��� $��1��$��1��$ 
Horizontal=�
�������$��2q�.+��� $�� ��$��Segoe)�$��16(�$
��Bold"�
���  ��$��Bottom=��� $��2��$��1��$ 
Horizontal=�
�������$��2q�.+��� -}   $ lbl_x_MonatAbw2��$
��+285��$��Segoe)�$��16(�$
��Bold"�
��� � �$��Bottom=��� -~   $��stp_jahr��$
��Left=�$��20 6 2 2q�.+��� -   $ 
lbl_x_Jahr��$
��Jahr��$��Segoe)�$��16(�$
��Bold"�
���   �
�������$��Center=�$��Center=�$��2q�� $	��160�� �� �� ��  �� �� .+��� $��0��$��0��$ 
Horizontal=�
�������$��2q�.+��� $	��✔��$��Segoe UI Symbol)�$��16(�$
��Bold"�
��� � ��� -�   $ lbl_x_JahrIst��$��-5 0 0 0q�$��10984��$��Segoe)�$��16(�$
��Bold"�
��� � �$��Bottom=��� $��1��$��0��$ 
Horizontal=�
�������$��2q�.+��� $��Ø��$��Segoe UI Symbol)�$��16(�$
��Bold"�
����  ��� -�   $ lbl_x_JahrSchnitt��$��-5 0 0 0q�$��40056��$��Segoe)�$��16(�$
��Bold"�
����  �$��Bottom=��� $��2��$��0��$ 
Horizontal=�
�������$��2q�.+��� $	��⇗��$��Segoe UI Symbol)�$��16(�$
��Bold"�
�������� -�   $ lbl_x_JahrHR��$��-5 0 0 0q�$��33229��$��Segoe)�$��16(�$
��Bold"�
������$��Bottom=��� $��0��$��1��$ 
Horizontal=�
�������$��2q�.+��� -�   $ lbl_x_JahrAbw1��$��-29072��$��Segoe)�$��16(�$
��Bold"�
���  ��$��Bottom=��� $��1��$��1��$ 
Horizontal=�
�������$��2q�.+��� $�� ��$��Segoe)�$��20(�$
��Bold"�
���  ��$��Bottom=��� $��2��$��1��$ 
Horizontal=�
�������$��2q�.+��� -�   $ lbl_x_JahrAbw2��$��-6736��$��Segoe)�$��16(�$
��Bold"�
���  ��$��Bottom=��� $ 
Horizontal=�.+��� -�    ��Name$ myGraph��$	��460��$	��320��
������ $��60 0 2 0q��� .+��� -�   $��
bt_RGChart��$��Chart: Abgerechnet��
�������$
��bold"�$��15(�$��60 14 14 20q�&   System.Windows.Controls.Calendar  $��	26 30 0 0q��� -�   $ ti_Daten��$��Daten��.���� $
��left=�$��20 2 2 2q�.+��� $ 
Horizontal=�$
��Left=�$��0q�.+��� -�   $ 	tb_Daten1��  �� 
AcceptsTab  .  . ! �� TextWrapping$
! Wrap=�$	��590��$	��460��$
 Auto=�$��16(�$��Consolas)��� $��20 0 0 0q�.+��� $��F2 = suchen��$��0 -6 0 0q�$��Black"�
�������� -�   $��tb_suche��$��0 -4 0 0q�$	��150��$�� ��
������$��Black"�$��16(�$��Consolas)��� $��F3 = Ersetzen��$��0 -5 0 0q�$��Black"��� -�   $��
tb_ersetze��$��0 -4 0 0q�$	��150��$�� ��
�������$��Black"�$��16(�$��Consolas)��� -�   $��bt_alleErsetzen��$��Alle Ersetzen��$��2 0 2 2q�$��Black"��� -�   $��bt_cutAtCurrent��$��Zeilenende abtrennen��$��2 8 2 2q�$��Black"��� -�   $��bt_cutFromBeginn��$��Zeilenanfang abtrennen��$��2 0 2 2q�$��Black"��� $��ab Cursorposition��$��6 -8 0 0q�$��Black"��� -�   $��bt_cutBlock��$��Markierten Block löschen��$��2 4 2 1q�$��Black"��� -�   $��bt_cutLeerzeilen��$��Leerzeilen löschen��$��2 0 2 1q�$��Black"��� -�   $��bt_F2ZeilenLoeschen��$��Zeilen wie F2 löschen��$��2 0 2 2q�$��Black"��� $ 
Horizontal=�.+��� .+��� -�   $��
bt_ladeTag��$��Daten: Tage��$��2 8 2 1q�$��Black"��� -�   $��	bt_ladeKW��$��Wochen��$��2 1 2 1q�$��Black"��� -�   $��bt_ladeProgPfade��$��	ProgPfade��$��2 2 2 1q�$��Black"��� $��75��.+��� -�   $��bt_ladeMonat��$��Monate��$��60��$��0 8 0 1q�$��Black"�$��right=��� -�   $��bt_ladeJahr��$��Jahre��$��60��$��0 1 0 0q�$��Black"�$��right=��� -�   $��bt_speichernProgPfade��$��
savePPfade��$��2 3 0 1q�$��Black"��� -�   $��lbl_textInfo��$��2 6 2 2q�.��Zeilen / Zeichen�� $ 
Horizontal=�.+��� $	��120��.+��� -�   $ tbwort1��$�� 0 0 2 0q�$	��120��$��20��$��SAP-Berater��$
��Left=��� -�   $ 
tbwort1alt��$�� 0 0 2 0q�$	��120��$��20��$��SAP-Consultant��$
��Left=��� -�   $ tbwort2��$�� 0 0 2 0q�$	��120��$��20��$��	WoTabZahl���� -�   $ 
tbwort2alt��$�� 0 0 2 0q�$	��120��$��20��$��TabZahlReturn���� -�   $��	bt_goText��$��Go��$��32��$��2 0 0 0q�$ 	0.502,0.5&��� $ 
Horizontal=�$��0 0 0 0q�.+�	� -�   $��
rb_staedte��$
 trueY�$��Städte��	� -�   $��
rb_laender��$��8 0 0 0q�$ falseY�$��Bundesländer���� -�   $ 	ti_Grafik��$��Grafik��.���� ���� $	 zwergenhaeuser.jpg��.+��� $ 
Horizontal=�.+��� .+��� -�   $��bt_clrChart��$
��Left=�$��Clear��
�������$
��bold"�$��15(�$��60��$��32��$	��Top=�$��10q��� -�   $��tb_charttitel��$��	10 10 0 2q�$	��120��$��Überschrift���� -�   $��	tb_xachse��$��10 0 0 2q�$	��120��$��X-Achsenbeschriftung���� -�   $��	tb_yachse��$��	10 0 0 10q�$	��120��$��Y-Achsenbeschriftung���� -�   $��bt_comChart��$
��Left=�$��Chart��
�������$
��bold"�$��15(�$��60��$��32��$	��Top=�$��10q��� -�   $��bt_comDaten��$
��Left=�$��Daten��
�������$
��bold"�$��15(�$��60��$��32��$	��Top=�$��10q��� $ 
Horizontal=�.+��� -�   $��bt_Canvas_kl��$
��Left=�$��--��
�������$
��bold"�$��15(�$��30��$��32��$	��Top=�$��10q��� -�   $��bt_Canvas_gr��$
��Left=�$��++��
�������$
��bold"�$��15(�$��30��$��32��$	��Top=�$��10q��� -�   $ 
myComGraph��$	��640��$	��470��
������ $��12 6 2 2q��� -�   $ 
ti_Termine��$��Termine��.���� $ 
Horizontal=�$
��Left=�.+��� -�   $ tb_Termine1��$��1 0 4 0q�  .  . $
! Wrap=�$	��470��$	��500��$
 Auto=�$" �� HorizontalScrollBarVisibility$
" Auto=�$��16(�$��Consolas)�$��tb_Termine1���� .+�� $	��120��$	��280��.+��� -�   $
��Btp1��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
���  ��$ Firefox��.���� .+��� -�   $��img14��$��32��$��32���� -�   $
��Btp2��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
���  ��$ Chrome��.���� .+��� -�   $��img15��$��32��$��32���� -�   $
��Btp3��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
���  ��$ Internet Explorer��.���� .+��� -�   $��img16��$��32��$��32���� -�   $
��Btp4��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
���  ��$ Excel��.���� .+��� -�   $��img17��$��32��$��32���� -�   $
��Btp5��$��5 5 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
���  ��$
 Word��.���� .+��� -�   $��img18��$��32��$��32���� -�   $
��Btp6��$��5 5 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
���  ��$ Indesign��.���� .+��� -�   $��img19��$��32��$��32���� -�   $
��Btp7��$��5 5 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
���  ��$  ��.���� .+��� -�   $��img20��$��32��$��32���� -�   $ tb_Termine2��  .  . $
! Wrap=�$	��350��$	��280��$
��Left=�$
 Auto=�$
" Auto=�$��16(�$��Consolas)�$��tb_Termine2���� 
���   �$��0 -5 0 0q�$ 
Horizontal=�$
��left=�$	��800��$��57��.+��� -�   $
��But1��$��12 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
���  ��$ 	IrfanView��.���� .+��� -�   $
��img1��$��32��$��32���� -�   $
��But2��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
���  ��$ 	Paint.NET��.���� .+��� -�   $
��img2��$��32��$��32���� -�   $
��But3��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
������$ Visual Studiio��.���� .+��� -�   $
��img3��$��32��$��32���� -�   $
��But4��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
���   �$ CyNote��.���� .+��� -�   $
��img4��$��32��$��32���� -�   $
��But5��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
����  �$ dnpNewsEditor��.���� .+��� -�   $
��img5��$��32��$��32���� -�   $
��But6��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
����  �$ 	NotePad++��.���� .+��� -�   $
��img6��$��32��$��32���� -�   $
��But7��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
����  �$ WriteMonkey��.���� .+��� -�   $
��img7��$��32��$��32���� -�   $
��But8��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
���  ��$ dotnetpro-Ordner��$	��dnp���� -�   $
��But9��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
��� � �$ Buchhaltung XLS��$
��BhXL���� -�   $��But10��$	��BhO��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
��� � �$ Ordner Buchhaltung���� -�   $��But11��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
��� � �$ Neue Rechnung XLS��$	��nRG���� -�   $��But12��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
���  ��$ HTML-Tabellen��.���� .+��� -�   $��img12��$��32��$��32���� -�   $��But13��$��5 0 5 0q�$��50��$��50��$��Consolas)�$��
ExtraBlack"�$��20(�
���  ��$ Bücherliste��.���� .+��� -�   $��img13��$��32��$��32��