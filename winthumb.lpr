library winthumb;

{$mode delphi}{$H+}
{$include calling.inc}

uses
  Classes,
  sysutils,
  WLXPlugin, Windows, ActiveX, ShlObj, Types;

const
  SIIGBF_RESIZETOFIT   = $00000000;
  SIIGBF_BIGGERSIZEOK  = $00000001;
  SIIGBF_MEMORYONLY    = $00000002;
  SIIGBF_ICONONLY      = $00000004;
  SIIGBF_THUMBNAILONLY = $00000008;
  SIIGBF_INCACHEONLY   = $00000010;

const
  IID_IExtractImage: TGUID = '{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}';

type
  SIIGBF = Integer;

  IShellItemImageFactory = interface(IUnknown)
    ['{BCC18B79-BA16-442F-80C4-8A59C30C463B}']
    function GetImage(size: TSize; flags: SIIGBF; out phbm: HBITMAP): HRESULT; stdcall;
  end;

  IExtractImage = interface(IUnknown)
    ['{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}']
    function GetLocation(pszPathBuffer: LPWSTR; cchMax: DWORD; out pdwPriority: DWORD;
      const prgSize: LPSIZE; dwRecClrDepth: DWORD; var pdwFlags: DWORD): HRESULT; stdcall;
    function Extract(out phBmpImage: HBITMAP): HRESULT; stdcall;
  end;

var
  SHCreateItemFromParsingName: function(pszPath: LPCWSTR; const pbc: IBindCtx;
                                        const riid: TIID; out ppv): HRESULT; stdcall;

procedure ListGetDetectString(DetectString:pchar;maxlen:integer); dcpcall;
begin
  StrCopy(DetectString, 'EXT="*"');
end;

function GetThumbnailOld(const aFileName: UTF8String; aSize: TSize; out Bitmap: HBITMAP): HRESULT;
var
  Folder,
  DesktopFolder: IShellFolder;
  Pidl,
  ParentPidl: PItemIDList;
  Image: IExtractImage;
  pchEaten: ULONG;
  wsTemp: WideString;
  dwPriority: DWORD;
  Status: HRESULT;
  dwRecClrDepth: DWORD;
  dwAttributes: ULONG = 0;
  dwFlags: DWORD = IEIFLAG_SCREEN or IEIFLAG_QUALITY or IEIFLAG_ORIGSIZE;
begin
  Result:= E_FAIL;

  if SHGetDesktopFolder(DesktopFolder) = S_OK then
  begin
    wsTemp:= UTF8Decode(ExtractFilePath(aFileName));
    if DesktopFolder.ParseDisplayName(0, nil, PWideChar(wsTemp), pchEaten, ParentPidl, dwAttributes) = S_OK then
    begin
      if DesktopFolder.BindToObject(ParentPidl, nil, IID_IShellFolder, Folder) = S_OK then
      begin
        wsTemp:= UTF8Decode(ExtractFileName(aFileName));
        if Folder.ParseDisplayName(0, nil, PWideChar(wsTemp), pchEaten, Pidl, dwAttributes) = S_OK then
        begin
          if Succeeded(Folder.GetUIObjectOf(0, 1, Pidl, IID_IExtractImage, nil, Image)) then
          begin
            SetLength(wsTemp, MAX_PATH * SizeOf(WideChar));
            dwRecClrDepth:= GetDeviceCaps(0, BITSPIXEL);
            Status:= Image.GetLocation(PWideChar(wsTemp), Length(wsTemp), dwPriority, @aSize, dwRecClrDepth, dwFlags);
            if (Status = NOERROR) or (Status = E_PENDING) then
            begin
              Result:= Image.Extract(Bitmap);
            end;
          end;
          CoTaskMemFree(Pidl);
        end;
        Folder:= nil;
      end;
      CoTaskMemFree(ParentPidl);
    end;
    DesktopFolder:= nil;
  end; // SHGetDesktopFolder
end;

function GetThumbnailNew(const aFileName: UTF8String; aSize: TSize; out Bitmap: HBITMAP): HRESULT;
var
  ShellItemImage: IShellItemImageFactory;
begin
  Result:= SHCreateItemFromParsingName(PWideChar(UTF8Decode(aFileName)), nil,
                                       IShellItemImageFactory, ShellItemImage);
  if Succeeded(Result) then
  begin
    Result:= ShellItemImage.GetImage(aSize, SIIGBF_THUMBNAILONLY, Bitmap);
  end;
end;

function ListGetPreviewBitmapFile(FileToLoad:pchar;OutputPath:pchar;width,height:integer;
    contentbuf:pchar;contentbuflen:integer):pchar; dcpcall;
var
  Bitmap: HBITMAP;
  Status: HRESULT = E_FAIL;
  headerSize : Size_t;
  pHeader: Pointer;
  pbmi : LPBITMAPINFO;
  bmf : BITMAPFILEHEADER;
  pData: Pointer;
  bFile : file;
begin
  try
    Result:= '';

    Bitmap := 0;

    if (Win32MajorVersion > 5) then
    begin
      Status:= GetThumbnailNew(FileToLoad, Size(width,height), Bitmap);
    end;

    if Failed(Status) then
    begin
      Status:= GetThumbnailOld(FileToLoad, Size(width,height), Bitmap);
    end;

    if Succeeded(Status) then
    begin
      headerSize := sizeof(BITMAPINFOHEADER)+3*sizeof(RGBQUAD);
      pHeader := AllocMem(headerSize);
      pbmi := LPBITMAPINFO(pHeader);
      FillChar(pHeader^,headerSize,0);
      pbmi^.bmiHeader.biSize := sizeof(BITMAPINFOHEADER);
      pbmi^.bmiHeader.biBitCount := 0;

      if (pbmi^.bmiHeader.biSizeImage <= 0) then
        pbmi^.bmiHeader.biSizeImage:=trunc(pbmi^.bmiHeader.biWidth*abs(pbmi^.bmiHeader.biHeight)*(pbmi^.bmiHeader.biBitCount+7)/8);
      pData := AllocMem(pbmi^.bmiHeader.biSizeImage);
      bmf.bfType := $4D42;
      bmf.bfReserved1 := 0;
      bmf.bfReserved2 := 0;
      bmf.bfSize := sizeof(BITMAPFILEHEADER)+ headerSize + pbmi^.bmiHeader.biSizeImage;
      bmf.bfOffBits := sizeof(BITMAPFILEHEADER) + headerSize;
      AssignFile(bFile,OutputPath+'thumb.bmp');
      Rewrite(bFile);
      BlockWrite(bFile,bmf,sizeof(BITMAPFILEHEADER));
      BlockWrite(bFile,pbmi^,headerSize);
      BlockWrite(bFile,pData^,pbmi^.bmiHeader.biSizeImage);
      Close(bFile);
      Freemem(pHeader);
      Freemem(pData);
      result := PChar(OutputPath+'thumb.bmp');
    end;
  except
    Result := '';
  end;
end;

exports
  ListGetDetectString,
  ListGetPreviewBitmapFile;

begin
  SHCreateItemFromParsingName := GetProcAddress(GetModuleHandle('shell32.dll'),
                                               'SHCreateItemFromParsingName');
end.

