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
    Width =10801
    DatasheetFontHeight =11
    ItemSuffix =13
    Left =540
    Top =533
    Right =11595
    Bottom =5213
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =4689
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
                    Width =10650
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
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =570
                    BackShade =75.0
                End
                Begin Subform
                    OverlapFlags =85
                    Left =143
                    Top =863
                    Width =10658
                    Height =3601
                    BorderColor =10921638
                    Name ="Form Resizer Example from pdtech.co.uk"
                    SourceObject ="Form.frmExampleWithTab"
                    Tag ="GrowX GrowY"
                    EventProcPrefix ="Form_Resizer_Example_from_pdtech_co_uk"
                    GridlineColor =10921638

                    LayoutCachedLeft =143
                    LayoutCachedTop =863
                    LayoutCachedWidth =10801
                    LayoutCachedHeight =4464
                End
            End
        End
    End
End
CodeBehindForm
' See "frmExampleWithSubform.cls"
