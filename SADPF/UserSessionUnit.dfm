object IWUserSession: TIWUserSession
  OldCreateOrder = False
  Height = 150
  Width = 215
  object ZQuery2: TZQuery
    Connection = ZConnection2
    Params = <>
    Left = 76
    Top = 24
  end
  object ZConnection2: TZConnection
    ControlsCodePage = cCP_UTF16
    HostName = 'localhost'
    Port = 3306
    Database = 'sadpf'
    User = 'root'
    Password = ''
    Protocol = 'mysql'
    Left = 76
    Top = 56
  end
end
