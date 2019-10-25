unit email;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IdMessage, IdTCPConnection, IdTCPClient, IdMessageClient,
  IdSMTP, IdBaseComponent, IdComponent, IdIOHandler, IdIOHandlerSocket,
  IdSSLOpenSSL, ComCtrls, Buttons, StdCtrls;

type
  TForm1 = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    ComboBox1: TComboBox;
    GroupBox2: TGroupBox;
    Label4: TLabel;
    Edit4: TEdit;
    Label5: TLabel;
    Edit5: TEdit;
    Label6: TLabel;
    Edit6: TEdit;
    GroupBox3: TGroupBox;
    Label7: TLabel;
    Label8: TLabel;
    Edit7: TEdit;
    Edit8: TEdit;
    Label9: TLabel;
    Label10: TLabel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    StatusBar1: TStatusBar;
    OpenDialog1: TOpenDialog;
    IdSSLIOHandlerSocket1: TIdSSLIOHandlerSocket;
    IdSMTP1: TIdSMTP;
    IdMessage1: TIdMessage;
    memo1: TMemo;
    ListBox1: TListBox;
    Label11: TLabel;
    Edit9: TEdit;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.SpeedButton1Click(Sender: TObject);
var
  i: Integer;
  IdMessage: TIdMessage;
begin
  if OpenDialog1.Execute then
    begin
      for i:=0 to OpenDialog1.Files.Count -1 do
        if (ListBox1.Items.IndexOf(OpenDialog1.Files[i])=-1) then
        ListBox1.Items.Add(OpenDialog1.Files[i]);
    end;
end;

procedure TForm1.SpeedButton2Click(Sender: TObject);
var
//objetos necessarios para o funcionamento
IdSSLIOHandlerSocket: TIdSSLIOHandlerSocket;
IdSMTP: TIdSMTP;
IdMessage: TIdMessage;
CaminhoAnexo: string;
i: Integer;
begin
//instanciação dos objetos
  IdSSLIOHandlerSocket:= TIdSSLIOHandlerSocket.Create(Self);
  IdSMTP:= TIdSMTP.Create(Self);
  IdMessage:= TIdMessage.Create(Self);
  try
  //configuração do SSL
  IdSSLIOHandlerSocket.SSLOptions.Method:=sslvSSLv23;
  IdSSLIOHandlerSocket.SSLOptions.Mode:=sslmClient;
  //Configuração do SMTP
  IdSMTP.IOHandler:=IdSSLIOHandlerSocket;
  IdSMTP.AuthenticationType:= atLogin;
  IdSMTP.Port:= StrToInt(ComboBox1.Text);
  IdSMTP.Host:= Edit1.Text;
  IdSMTP.Username:= Edit2.Text;
  IdSMTP.Password:=Edit3.Text;
  //tentativa de conexao e autenticação
  try
    IdSMTP.Connect;
    IdSMTP.Authenticate;
    except
    on E:Exception do
      begin
        MessageDlg('Erro na conexão e/ou autenticação: '
                    + E.Message, mtWarning, [mbOK], 0);
        Exit;
      end;
    end;
//Configuração da mensagem
  IdMessage.From.Address:= Edit8.Text;
  IdMessage.From.Name := Edit9.Text;
  IdMessage.ReplyTo.EMailAddresses:= IdMessage.From.Address;
  IdMessage.Recipients.EMailAddresses:=Edit4.Text;
  IdMessage.CCList.EMailAddresses:=Edit5.Text;
  IdMessage.BccList.EMailAddresses:=Edit6.Text;
  IdMessage.Subject:= Edit7.Text;
  IdMessage.Body.Text:=memo1.Lines.Text;
  //Anexo da mensagem (opcional)
  if ListBox1.Items.Count > 0 then
    begin
      for i:= 0 to ListBox1.Items.Count -1 do
        begin
          if FileExists(ListBox1.Items [i]) then
          TIdAttachment.Create(IdMessage.MessageParts, ListBox1.Items[i]);
        end;
    end;
//Envio da mensagem
  try
    IdSMTP.Send(IdMessage);
    MessageDlg('mensagem enviada com sucesso.', mtInformation, [mbOK], 0);
    except
      on E:Exception do
      MessageDlg('Erro ao enviar a mensagem: ' + E.Message, mtWarning, [mbOK], 0);
  end;
  finally
//liberação dos objetos da memória
  FreeAndNil(IdMessage);
  FreeAndNil(IdSSLIOHandlerSocket);
  FreeAndNil(IdSMTP);
  end;

end;

end.
