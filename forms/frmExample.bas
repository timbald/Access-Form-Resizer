Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7056
    DatasheetFontHeight =11
    ItemSuffix =14
    Left =645
    Top =765
    Right =10560
    Bottom =6720
    RecSrcDt = Begin
        0x4acbaa143b87e440
    End
    Caption ="Form Resizer Example from pdtech.co.uk"
    DatasheetFontName ="Calibri"
    OnResize ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =4981
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =60.0
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =150
                    Top =60
                    Width =6675
                    Height =510
                    FontSize =18
                    BackColor =12566463
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label0"
                    Caption ="Resize options set via Tag property"
                    Tag ="GrowX"
                    GridlineColor =10921638
                    LayoutCachedLeft =150
                    LayoutCachedTop =60
                    LayoutCachedWidth =6825
                    LayoutCachedHeight =570
                    BackShade =75.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =165
                    Top =1342
                    Width =6642
                    Height =675
                    FontSize =26
                    BackColor =12566463
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label1"
                    Caption ="GrowX GrowY0.5 ScaleFont"
                    OnDblClick ="[Event Procedure]"
                    Tag ="GrowY0.5 GrowX ScaleFont"
                    GridlineColor =10921638
                    LayoutCachedLeft =165
                    LayoutCachedTop =1342
                    LayoutCachedWidth =6807
                    LayoutCachedHeight =2017
                    BackShade =75.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =165
                    Top =2160
                    Width =3285
                    Height =2190
                    FontSize =18
                    BackColor =12566463
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label2"
                    Caption ="GrowX0.5 GrowY0.5 MoveY0.5\015\012https://github.com/pd-oflaherty/Access-Form-Re"
                        "sizer"
                    Tag ="GrowX0.5 GrowY0.5 MoveY0.5 ScaleFont"
                    GridlineColor =10921638
                    LayoutCachedLeft =165
                    LayoutCachedTop =2160
                    LayoutCachedWidth =3450
                    LayoutCachedHeight =4350
                    BackShade =75.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =165
                    Top =701
                    Width =1643
                    Height =510
                    FontSize =18
                    BackColor =12566463
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label3"
                    Caption ="GrowX0.25"
                    Tag ="GrowX0.25"
                    GridlineColor =10921638
                    LayoutCachedLeft =165
                    LayoutCachedTop =701
                    LayoutCachedWidth =1808
                    LayoutCachedHeight =1211
                    BackShade =75.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =1947
                    Top =690
                    Width =1527
                    Height =510
                    FontSize =18
                    BackColor =12566463
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label4"
                    Caption ="GrowX0.25 MoveX0.25"
                    Tag ="GrowX0.25 MoveX0.25"
                    GridlineColor =10921638
                    LayoutCachedLeft =1947
                    LayoutCachedTop =690
                    LayoutCachedWidth =3474
                    LayoutCachedHeight =1200
                    BackShade =75.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =3613
                    Top =690
                    Width =1527
                    Height =510
                    FontSize =18
                    BackColor =12566463
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label5"
                    Caption ="GrowX0.25 MoveX0.5"
                    Tag ="GrowX0.25 MoveX0.5"
                    GridlineColor =10921638
                    LayoutCachedLeft =3613
                    LayoutCachedTop =690
                    LayoutCachedWidth =5140
                    LayoutCachedHeight =1200
                    BackShade =75.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =5279
                    Top =690
                    Width =1527
                    Height =510
                    FontSize =18
                    BackColor =12566463
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label6"
                    Caption ="GrowX0.25 MoveX0.75"
                    Tag ="GrowX0.25 MoveX0.75"
                    GridlineColor =10921638
                    LayoutCachedLeft =5279
                    LayoutCachedTop =690
                    LayoutCachedWidth =6806
                    LayoutCachedHeight =1200
                    BackShade =75.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =150
                    Top =4470
                    Width =1643
                    Height =510
                    FontSize =18
                    BackColor =12566463
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label7"
                    Caption ="MoveY"
                    Tag ="MoveY"
                    GridlineColor =10921638
                    LayoutCachedLeft =150
                    LayoutCachedTop =4470
                    LayoutCachedWidth =1793
                    LayoutCachedHeight =4980
                    BackShade =75.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =1932
                    Top =4471
                    Width =1527
                    Height =510
                    FontSize =18
                    BackColor =12566463
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label8"
                    Caption ="GrowX0.5 MoveY"
                    Tag ="GrowX0.5 MoveY"
                    GridlineColor =10921638
                    LayoutCachedLeft =1932
                    LayoutCachedTop =4471
                    LayoutCachedWidth =3459
                    LayoutCachedHeight =4981
                    BackShade =75.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =3598
                    Top =4471
                    Width =1527
                    Height =510
                    FontSize =18
                    BackColor =12566463
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label9"
                    Caption ="GrowX0.5 MoveX0.5 MoveY"
                    Tag ="GrowX0.5 MoveX0.5 MoveY"
                    GridlineColor =10921638
                    LayoutCachedLeft =3598
                    LayoutCachedTop =4471
                    LayoutCachedWidth =5125
                    LayoutCachedHeight =4981
                    BackShade =75.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =5264
                    Top =4471
                    Width =1527
                    Height =510
                    FontSize =18
                    BackColor =12566463
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label10"
                    Caption ="MoveX MoveY"
                    Tag ="MoveX MoveY"
                    GridlineColor =10921638
                    LayoutCachedLeft =5264
                    LayoutCachedTop =4471
                    LayoutCachedWidth =6791
                    LayoutCachedHeight =4981
                    BackShade =75.0
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3594
                    Top =2160
                    Width =3174
                    Height =2190
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text12"
                    DefaultValue ="\"Enter some example text here, then resize to see the effect of the ScaleFont t"
                        "ag. Font will change but the text will not be vertically aligned (as is the case"
                        " with labels). Tag property for this control is: MoveX0.5 GrowX0.5 GrowY0.5 Move"
                        "Y0.5 ScaleFont\""
                    Tag ="MoveX0.5 GrowX0.5 GrowY0.5 MoveY0.5 ScaleFont"
                    GridlineColor =10921638

                    LayoutCachedLeft =3594
                    LayoutCachedTop =2160
                    LayoutCachedWidth =6768
                    LayoutCachedHeight =4350
                End
            End
        End
    End
End
CodeBehindForm
' See "frmExample.cls"
