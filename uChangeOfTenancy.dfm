object FrmChangeOfTenancy: TFrmChangeOfTenancy
  AlignWithMargins = True
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  BorderStyle = bsNone
  BorderWidth = 1
  Caption = 'FrmChangeOfTenancy'
  ClientHeight = 619
  ClientWidth = 721
  Color = clBtnFace
  DockSite = True
  DragMode = dmAutomatic
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object pnlBack: TPanel
    Left = 0
    Top = 0
    Width = 721
    Height = 619
    Align = alClient
    BevelEdges = [beLeft, beTop, beBottom]
    BevelOuter = bvNone
    Color = clWhite
    ParentBackground = False
    TabOrder = 0
    object shpTopBar: TShape
      Left = 0
      Top = 0
      Width = 721
      Height = 37
      Align = alTop
      Brush.Color = clHotLight
      Pen.Color = clHotLight
      ExplicitTop = 8
      ExplicitWidth = 546
    end
    object lblTopBar: TLabel
      Left = 22
      Top = 10
      Width = 123
      Height = 16
      Caption = 'Change of Tenancy'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clHighlightText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblClose: TLabel
      Left = 684
      Top = 6
      Width = 25
      Height = 25
      Cursor = crHandPoint
      Alignment = taCenter
      AutoSize = False
      Caption = 'X'
      Font.Charset = ANSI_CHARSET
      Font.Color = 16119285
      Font.Height = -16
      Font.Name = 'Default'
      Font.Style = [fsBold]
      ParentFont = False
      Layout = tlCenter
      OnClick = btnCloseClick
    end
    object pgcCustomer: TPageControl
      AlignWithMargins = True
      Left = 5
      Top = 82
      Width = 711
      Height = 494
      Margins.Left = 5
      Margins.Right = 5
      ActivePage = tsOutgoingCustomer
      Align = alBottom
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      Images = DM_Images.LargeImages
      ParentFont = False
      TabOrder = 0
      object tsOutgoingCustomer: TTabSheet
        Caption = 'Outgoing Customer'
        ImageIndex = 137
        object pgcAgreement: TPageControl
          AlignWithMargins = True
          Left = 3
          Top = 63
          Width = 697
          Height = 383
          Align = alBottom
          Images = DM_Images.LargeImages
          TabOrder = 0
        end
        object pnlOutgoingCustomerInfo: TPanel
          AlignWithMargins = True
          Left = 3
          Top = 3
          Width = 697
          Height = 54
          Align = alTop
          Alignment = taLeftJustify
          BevelOuter = bvNone
          TabOrder = 1
          object shpCustomerName: TShape
            Left = 3
            Top = 22
            Width = 222
            Height = 27
            Pen.Color = clHotLight
            Shape = stRoundRect
          end
          object lblCustomerName: TLabel
            Left = 3
            Top = 3
            Width = 76
            Height = 13
            Caption = 'Customer Name'
            Color = clHotLight
            ParentColor = False
          end
          object shpForwardingAddress: TShape
            Left = 520
            Top = 20
            Width = 169
            Height = 27
            Brush.Color = clHotLight
            Pen.Color = clHotLight
            Shape = stRoundRect
          end
          object lblForwardingAddress: TLabel
            Left = 520
            Top = 22
            Width = 169
            Height = 25
            Cursor = crHandPoint
            Alignment = taCenter
            AutoSize = False
            Caption = 'Forwarding Address'
            Font.Charset = ANSI_CHARSET
            Font.Color = 16119285
            Font.Height = -12
            Font.Name = 'Urbanist'
            Font.Style = [fsBold]
            ParentFont = False
            Layout = tlCenter
            OnClick = lblForwardingAddressClick
          end
          object lblOutCustWalletBalance: TLabel
            Left = 250
            Top = 3
            Width = 70
            Height = 13
            Caption = 'Wallet Balance'
            Color = clHotLight
            ParentColor = False
          end
          object shpOutCustWalletBalance: TShape
            Left = 250
            Top = 22
            Width = 222
            Height = 27
            Pen.Color = clHotLight
            Shape = stRoundRect
          end
          object btnForwardingAddress: TButton
            Left = 631
            Top = -1
            Width = 41
            Height = 25
            Caption = 'Forwarding Address'
            TabOrder = 0
            Visible = False
          end
          object edtCustomerName: TEdit
            Left = 10
            Top = 28
            Width = 207
            Height = 19
            BorderStyle = bsNone
            TabOrder = 1
          end
          object edtOutCustWalletBalance: TEdit
            Left = 255
            Top = 28
            Width = 207
            Height = 19
            BorderStyle = bsNone
            TabOrder = 2
          end
        end
      end
      object tsIncomingCustomer: TTabSheet
        Caption = 'Incoming Customer'
        ImageIndex = 138
        object pgcAgreementIn: TPageControl
          AlignWithMargins = True
          Left = 3
          Top = 119
          Width = 697
          Height = 327
          Align = alBottom
          TabOrder = 0
          ExplicitWidth = 652
        end
        object pnlCustomerDetails: TPanel
          Left = 0
          Top = 0
          Width = 703
          Height = 113
          Align = alTop
          BevelOuter = bvNone
          TabOrder = 1
          object lblInCustomerName1: TLabel
            Left = 4
            Top = 3
            Width = 110
            Height = 14
            Caption = 'Customer Forename'
            Color = clHotLight
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object shpInCustomerName1: TShape
            Left = 15
            Top = 17
            Width = 138
            Height = 27
            Pen.Color = clHotLight
            Shape = stRoundRect
          end
          object lblInCustomerDoB1: TLabel
            Left = 335
            Top = 3
            Width = 70
            Height = 14
            Caption = 'Date of Birth'
            Color = clHotLight
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object shpInCustomerDoB: TShape
            Left = 348
            Top = 17
            Width = 128
            Height = 27
            Pen.Color = clHotLight
            Shape = stRoundRect
          end
          object lblInCustomerEmail: TLabel
            Left = 4
            Top = 51
            Width = 30
            Height = 14
            Caption = 'e-Mail'
            Color = clHotLight
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object shpInCustomerEmail: TShape
            Left = 15
            Top = 66
            Width = 295
            Height = 27
            Pen.Color = clHotLight
            Shape = stRoundRect
          end
          object lblInCustomerMobile: TLabel
            Left = 335
            Top = 51
            Width = 107
            Height = 14
            Caption = 'Telephone (Mobile)'
            Color = clHotLight
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object shpInCustomerMobileA: TShape
            Left = 348
            Top = 66
            Width = 41
            Height = 27
            Brush.Color = clMenuHighlight
            Pen.Color = clHotLight
            Shape = stRoundRect
          end
          object lblInCustomerName2: TLabel
            Left = 169
            Top = 3
            Width = 48
            Height = 14
            Caption = 'Surname'
            Color = clHotLight
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object shpInCustomerName2: TShape
            Left = 172
            Top = 17
            Width = 138
            Height = 27
            Pen.Color = clHotLight
            Shape = stRoundRect
          end
          object shpInCustomerMobileB: TShape
            Left = 395
            Top = 66
            Width = 47
            Height = 27
            Pen.Color = clHotLight
            Shape = stRoundRect
          end
          object shpInCustomerMobileC: TShape
            Left = 450
            Top = 66
            Width = 79
            Height = 27
            Pen.Color = clHotLight
            Shape = stRoundRect
          end
          object lblInCustomerDoB2: TLabel
            Left = 425
            Top = 7
            Width = 48
            Height = 10
            Caption = 'dd-mm-YYYY'
            Color = clHotLight
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -8
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object edtInCustomerMobileInfoA: TLabel
            Left = 390
            Top = 74
            Width = 4
            Height = 13
            Caption = '-'
          end
          object edtInCustomerMobileInfoB: TLabel
            Left = 444
            Top = 73
            Width = 4
            Height = 13
            Caption = '-'
          end
          object edtInCustomerMobileInfoC: TLabel
            Left = 405
            Top = 92
            Width = 36
            Height = 10
            Caption = 'Area Code'
            Color = clHotLight
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -8
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object edtInCustomerMobileInfoD: TLabel
            Left = 479
            Top = 92
            Width = 50
            Height = 10
            Caption = 'Phone Number'
            Color = clHotLight
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -8
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object lblInCustomerContact: TLabel
            Left = 496
            Top = 3
            Width = 144
            Height = 14
            Caption = 'Preferred Contact Method'
            Color = clHotLight
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentColor = False
            ParentFont = False
          end
          object edtInCustomerName1: TEdit
            Left = 24
            Top = 23
            Width = 122
            Height = 19
            BorderStyle = bsNone
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            MaxLength = 35
            ParentFont = False
            TabOrder = 0
            OnKeyPress = edtInCustomerName1KeyPress
          end
          object edtInCustomerEmail: TEdit
            Left = 24
            Top = 72
            Width = 279
            Height = 19
            BorderStyle = bsNone
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 3
          end
          object edtInCustomerMobileA: TEdit
            Left = 353
            Top = 72
            Width = 30
            Height = 19
            TabStop = False
            BevelInner = bvNone
            BorderStyle = bsNone
            Color = clMenuHighlight
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWhite
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = [fsBold]
            ImeMode = imDisable
            ParentFont = False
            ReadOnly = True
            TabOrder = 7
            Text = '+44'
          end
          object DateInCustomerDoB: TDateTimePicker
            Left = 351
            Top = 19
            Width = 122
            Height = 22
            BevelInner = bvNone
            BevelOuter = bvNone
            Date = 2.000000000000000000
            Time = 2.000000000000000000
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            MaxDate = 40179.999988425920000000
            MinDate = 2.000000000000000000
            ParentFont = False
            TabOrder = 2
          end
          object cboInContactChoice: TComboBox
            Left = 512
            Top = 20
            Width = 138
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 6
            Items.Strings = (
              'e-Mail'
              'Mobile')
          end
          object edtInCustomerName2: TEdit
            Left = 181
            Top = 23
            Width = 122
            Height = 19
            BorderStyle = bsNone
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            MaxLength = 35
            ParentFont = False
            TabOrder = 1
            OnKeyPress = edtInCustomerName1KeyPress
          end
          object edtInCustomerMobileB: TEdit
            Left = 401
            Top = 72
            Width = 35
            Height = 19
            BorderStyle = bsNone
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            MaxLength = 5
            ParentFont = False
            TabOrder = 4
            OnKeyPress = edtInCustomerMobileBKeyPress
          end
          object edtInCustomerMobileC: TEdit
            Left = 456
            Top = 72
            Width = 69
            Height = 19
            BorderStyle = bsNone
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clBlack
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            MaxLength = 9
            ParentFont = False
            TabOrder = 5
            OnKeyPress = edtInCustomerMobileBKeyPress
          end
        end
      end
    end
    object pnlButtons: TPanel
      Left = 0
      Top = 579
      Width = 721
      Height = 40
      Align = alBottom
      BevelEdges = [beLeft, beTop, beBottom]
      BevelOuter = bvNone
      Color = clWhite
      ParentBackground = False
      TabOrder = 1
      object shpBtnCancel: TShape
        Left = 588
        Top = 3
        Width = 121
        Height = 27
        Brush.Color = clHotLight
        Pen.Color = clHotLight
        Shape = stRoundRect
      end
      object shpBtnSubmit: TShape
        Left = 438
        Top = 3
        Width = 121
        Height = 27
        Brush.Color = clHotLight
        Pen.Color = clHotLight
        Shape = stRoundRect
      end
      object lblSubmit: TLabel
        Left = 438
        Top = 3
        Width = 121
        Height = 25
        Cursor = crHandPoint
        Alignment = taCenter
        AutoSize = False
        Caption = 'Submit'
        Font.Charset = ANSI_CHARSET
        Font.Color = 16119285
        Font.Height = -12
        Font.Name = 'Urbanist'
        Font.Style = [fsBold]
        ParentFont = False
        Layout = tlCenter
        OnClick = btnSubmitClick
      end
      object lblCancel: TLabel
        Left = 588
        Top = 3
        Width = 121
        Height = 25
        Cursor = crHandPoint
        Alignment = taCenter
        AutoSize = False
        Caption = 'Cancel'
        Font.Charset = ANSI_CHARSET
        Font.Color = 16119285
        Font.Height = -12
        Font.Name = 'Urbanist'
        Font.Style = [fsBold]
        ParentFont = False
        Layout = tlCenter
        OnClick = btnCloseClick
      end
      object pbPerformingCot: TProgressBar
        Left = 245
        Top = 8
        Width = 150
        Height = 17
        Style = pbstMarquee
        Step = 6
        TabOrder = 0
        Visible = False
      end
    end
    object pnlVacateDate: TPanel
      Left = 0
      Top = 37
      Width = 721
      Height = 36
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 2
      object lblVacatedDate: TLabel
        Left = 11
        Top = 11
        Width = 75
        Height = 14
        Caption = 'Vacated Date'
        Color = clHotLight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object DateTimeChangeTenant: TDateTimePicker
        Left = 108
        Top = 11
        Width = 186
        Height = 22
        Date = 2.000000000000000000
        Time = 0.500000000000000000
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
    end
  end
end