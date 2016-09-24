unit main;

interface
uses
  Windows, Classes, ActiveX, ShlObj, ComServ, ComObj,
  Urlmon, registry, axctrls, SysUtils,
  wininet, PythonEngine, VarPyth;

var Py : TPythonEngine;
var PyModule : Variant;

const
  MimeFilterType = 'text/html';
  MimeFilterName = 'mkrz ieSnake';
  CLSID_MimeFilter: TGUID = '{2FC29FB7-2BD4-450B-851E-89C56C86A635}';
  // kk made this GUID with Ctrl-Shift-G

type
  TMimeFilterFactory = class(TComObjectFactory)
  private
    procedure AddKeys;
    procedure RemoveKeys;
  public
    procedure UpdateRegistry(Register: Boolean); override;
  end;

type
  TMimeFilter = class(TComObject, IInternetProtocol, IInternetProtocolSink)
  private
    CacheFileName: string;
    Url: PWideChar;
    DataStream: IStream;
    UrlMonProtocol: IInternetProtocol;
    UrlMonProtocolSink: IInternetProtocolSink;
    Written, TotalSize: Integer;
  protected
    // IInternetProtocolSink Methods
    function Switch(const ProtocolData: TProtocolData): HResult; stdcall;
    function ReportProgress(ulStatusCode: ULONG; szStatusText: LPCWSTR): HResult; stdcall;
    function ReportData(grfBSCF: DWORD; ulProgress, ulProgressMax: ULONG): HResult; stdcall;
    function ReportResult(hrResult: HResult; dwError: DWORD; szResult: LPCWSTR): HResult; stdcall;
    // IInternetProtocol Methods
    function Start(szUrl: PWideChar; OIProtSink: IInternetProtocolSink; OIBindInfo: IInternetBindInfo; grfPI, dwReserved: DWORD): HResult; stdcall;
    function Continue(const ProtocolData: TProtocolData): HResult; stdcall;
    function Abort(hrReason: HResult; dwOptions: DWORD): HResult; stdcall;
    function Terminate(dwOptions: DWORD): HResult; stdcall;
    function Suspend: HResult; stdcall;
    function Resume: HResult; stdcall;
    function Read(pv: Pointer; cb: ULONG; out cbRead: ULONG): HResult; stdcall;
    function Seek(dlibMove: LARGE_INTEGER; dwOrigin: DWORD; out libNewPosition: ULARGE_INTEGER): HResult; stdcall;
    function LockRequest(dwOptions: DWORD): HResult; stdcall;
    function UnlockRequest: HResult; stdcall;
  end;

implementation

function TMimeFilter.Start(szUrl: PWideChar; OIProtSink: IInternetProtocolSink; OIBindInfo: IInternetBindInfo; grfPI, dwReserved: DWORD): HResult;
var
  Fetched: Cardinal;
begin
  CacheFileName := '';
  TotalSize := 0;
  Written := 0;
  UrlMonProtocol := OIProtSink as IInternetProtocol;
  UrlMonProtocolSink := OIProtSink as IInternetProtocolSink;
  OIBindinfo.GetBindString(BINDSTRING_URL, @Url, 1, Fetched);
  Result := S_OK;
end;

function TMimeFilter.ReportProgress(ulStatusCode: ULONG; szStatusText: LPCWSTR): HResult;
begin
  if ulStatusCode = BINDSTATUS_CACHEFILENAMEAVAILABLE then
    CacheFileName := SzStatusText;
  UrlMonProtocolSink.ReportProgress(ulStatusCode, szStatustext);
  Result := S_OK;
end;

function TMimeFilter.ReportData(grfBSCF: DWORD; ulProgress,
  ulProgressMax: ULONG): HResult;
var
  TS: TStringStream;
  Dummy: Int64;
  hr: HResult;
  readTotal: ULONG;
  S: string;
  Fname: array[0..512] of Char;
  p: array[0..1000] of char;
  i: integer;
begin

  Ts := TStringStream.Create('');
  repeat
    hr := UrlMonProtocol.Read(@P, SizeOf(p), Readtotal);
    Ts.write(P, Readtotal);
  until (hr = S_FALSE) or (hr = INET_E_DOWNLOAD_FAILURE) or (hr = INET_E_DATA_NOT_AVAILABLE);

  if hr = S_FALSE then begin
    if CacheFilename = '' then begin
      //codesite.send('CacheFilename is blank');
      CreateUrlCacheEntry(@url, ts.Size, Pchar('htm'), FName, 0);
      TMemoryStream(ts).SaveToFile(Fname);
      StringToWideChar(StrPas(FName), @FName, SizeOf(FName));
      ReportProgress(BINDSTATUS_CACHEFILENAMEAVAILABLE, @FName);
    end;

    //S := StringReplace(Ts.DataString, 'Delphi', 'Borland Inprise', [rfReplaceAll, rfIgnoreCase]);
    {
    1. (use ieSnake.pth or something)
    2. import some module
    3. call ieSnake(url,html) in that module
    4. return results of that call
    }
    // TODO : in the future, allow enable/disable of filter by url I guess
    try
      S := PyModule.hook(string(Url),Ts.DataString);
      //codesite.send('S',S);
    except
      on e:Exception do begin
        //codesite.send('Exception',e);
        S := '<b>' + e.Message + '</b>';
        for i:= Py.Traceback.ItemCount-1 downto 0 do begin
          S := S + '<br>' + Py.Traceback.Items[i].FileName + ' (' + inttostr(Py.Traceback.Items[i].LineNo) + ') : ' + Py.Traceback.Items[i].Context;
        end;
      end;
    end;

    ts.Size := 0;
    ts.WriteString(S);

    TotalSize := Ts.Size;
    ts.Seek(0, 0);
    CreateStreamOnHGlobal(0, True, DataStream);
    TOlestream.Create(DataStream).CopyFrom(ts, ts.size);
    TS.Free;
    DataStream.Seek(0, STREAM_SEEK_SET, Dummy);
    UrlMonProtocolSink.ReportData(BSCF_FIRSTDATANOTIFICATION or BSCF_LASTDATANOTIFICATION or BSCF_DATAFULLYAVAILABLE, TotalSize, Totalsize);
    UrlMonProtocolSink.ReportResult(S_OK, S_OK, nil);

  end else begin
    Abort(hr, 0); //On Error: INET_E_DOWNLOAD_FAILURE or INET_E_DATA_NOT_AVAILABLE
  end;
  Result := S_OK;
end;

function TMimeFilter.Read(pv: Pointer; cb: ULONG; out cbRead: ULONG): HResult;
begin
  DataStream.Read(pv, cb, @cbRead);
  Inc(written, cbread);
  if (written = totalsize) then result := S_FALSE else Result := S_OK;
end;

function TMimeFilter.Continue(const ProtocolData: TProtocolData): HResult;
begin
  UrlMonProtocol.Continue(ProtocolData);
  result := S_OK;
end;

function TMimeFilter.Terminate(dwOptions: DWORD): HResult;
begin
  UrlmonProtocol.Terminate(dwOptions);
  result := S_OK;
end;

function TMimeFilter.Abort(hrReason: HResult; dwOptions: DWORD): HResult;
begin
  UrlMonProtocol.Abort(hrReason, dwOptions);
  result := S_OK;
end;

function TMimeFilter.LockRequest(dwOptions: DWORD): HResult;
begin
  UrlMonProtocol.LockRequest(dwOptions);
  result := S_OK;
end;

function TMimeFilter.UnlockRequest: HResult;
begin
  UrlMonProtocol.UnlockRequest;
  result := S_OK;
end;

function TMimeFilter.Seek(dlibMove: LARGE_INTEGER; dwOrigin: DWORD;
  out libNewPosition: ULARGE_INTEGER): HResult;
begin
  UrlMonProtocol.Seek(dlibMove, dwOrigin, libNewPosition);
  result := S_OK;
end;

function TMimeFilter.Suspend: HResult;
begin
  result := E_NOTIMPL;
end;

function TMimeFilter.Resume: HResult;
begin
  result := E_NOTIMPL;
end;

function TMimeFilter.Switch(const ProtocolData: TProtocolData): HResult;
begin
  UrlMonProtocolSink.Switch(ProtocolData);
  result := S_OK;
end;

function TMimeFilter.ReportResult(hrResult: HResult; dwError: DWORD; szResult: LPCWSTR): HResult;
begin
//MRZ from MS C example
	hrResult := S_OK;
	if (UrlMonProtocolSink = nil) then begin
//    codesite.send('UrlMonProtocolSink is nil');
		result := E_FAIL;
// MRZ
  end else begin
//    codesite.send('UrlMonProtocolSink Reporting Begin');
    UrlMonProtocolSink.ReportResult(hrResult, dwError, szResult);
//    codesite.send('UrlMonProtocolSink Reporting End');
    Result := S_OK;
  end;
end;

procedure TMimeFilterFactory.UpdateRegistry(Register: Boolean);
begin
  inherited UpdateRegistry(Register);
  if Register then AddKeys else RemoveKeys;
end;

procedure TMimeFilterFactory.AddKeys;
var S: string;
begin
  S := GUIDToString(CLSID_MimeFilter);
  with TRegistry.Create do
  try
    RootKey := HKEY_CLASSES_ROOT;
    if OpenKey('PROTOCOLS\Filter\' + MimeFilterType, True) then
    begin
      WriteString('', MimeFilterName);
      WriteString('CLSID', S);
      CloseKey;
    end;
  finally
    Free;
  end;
end;

procedure TMimeFilterFactory.RemoveKeys;
var S: string;
begin
  S := GUIDToString(CLSID_MimeFilter);
  with TRegistry.Create do
  try
    RootKey := HKEY_CLASSES_ROOT;
    DeleteKey('PROTOCOLS\Filter\'+MimeFilterType );
  finally
    Free;
  end;
end;

initialization
  TMimeFilterFactory.Create(
    ComServer, TMimeFilter, CLSID_MimeFilter,
    'ieSnake', 'mkrz ieSnake', ciMultiInstance, tmApartment // MRZ adds tmApartment
  );
  Py := TPythonEngine.Create(nil);
  Py.LoadDLL();
  PyModule := Import('ieSnakeHook');
finalization
  Py.Free();
end.

