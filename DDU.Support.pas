Unit DDU.Support;

//*****************************************************************************
//
// DDULIBRARY (DDU.Support)
// Copyright 2020 Clinton R. Johnson (xepol@xepol.com)
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
// Version : 5.0
//
// Purpose :
//
// History : <none>
//
//*****************************************************************************

Interface

{$I DVer.inc}
{$Warn SYMBOL_PLATFORM off}
{$Warn SYMBOL_DEPRECATED OFF}

Uses
  WinAPI.Windows,
  WinAPI.ShellAPI,

  System.SysUtils,
  System.StrUtils,
  System.DateUtils,
  System.Win.Registry,
  System.Classes,
  System.TypInfo,
  System.Math,
  System.Masks,
  DDU.SupportUnit;


{$I DTypes.inc}

Const
  MillisecondsPerSecond             = (1000);
  MillisecondsPerMinute             = (60*MillisecondsPerSecond);
  MillisecondsPerHour               = (60*MillisecondsPerMinute);
  MillisecondsPerDay                = (24*MillisecondsPerHour);
  MillisecondsPerWeek               = (7*MillisecondsPerDay);

Type
  TDataEncoding =(deAnsi,deUTF8,deUTF16);

Const
  BooleanText           : Array[Boolean] of String = ('F','T');
  TPunctuation          : Set Of AnsiChar = [#32..#47,#58..#64,#91..#96,#123..#127];
  TInValidFileName      : Set Of AnsiChar = [#0..#31,'\','/',':','?','"','<','>','|'];
  MaxCardinalValue      : Cardinal = HIgh(Cardinal);
  MinCardinalValue      : Cardinal = Low(Cardinal);

Type
  TShellFolder =( sfAppData,sfCache,sfCookies,sfDesktop,sfFavorites,sfFonts,sfHistory,sfLocalAppData,sfLocalSettings,sfMyMusic,sfMyPictures,
                   sfMyVideo,sfNetHood,sfPersonal,sfPrintHood,sfPrograms,sfRecent,sfSendTo,sfStartMenu,sfStartup,sfTemplates);

Const
  ShellFolders : Array[TShellFolder] of String =('AppData','Cache','Cookies','Desktop','Favorites','Fonts',
                                                 'History','Local AppData','Local Settings','My Music','My Pictures',
                                                 'My Video','Nethood','Personal','PrintHood','Programs',
                                                 'Recent','SendTo','Start Menu','Startup','Templates');

Type
  TidOS=(idOSUnknown,idOSWin95,idOSWin98,idOSWin98SE,idOSWinME,idOSWinNT,idOSWin2000, idOSXP,idOSVista,idOSWin7,idOSWin8,idOSWin10 );
  TAnsiCharSet          = Set of AnsiChar;
  TCharSet              = Set Of AnsiChar;
  TInternalDebugMethod  = Procedure(What : String) Of Object;
  TInternalDebugProc    = Procedure (What : String);
  TOnWalkTreeEvent      = Procedure(FullPath : String; SearchRec : TSearchRec) Of Object;
  CopyTreeException     = Class(Exception) End;
  MoveTreeException     = Class(Exception) End;

  TERect                = Record
    Left,Bottom,Right,Top : Extended;
  End;


  TDDUSystem = Class
  Private
    fUserName        : String;
    fWorkstation     : String;
    fNetworkUserName : String;
    fSystemPath      : String;
    fTempPath        : String;
    fWindowsPath     : String;
    fIDOS            : TidOS;
    fProgramFilesPath: String;
    function GetOperatingSystem: TIDOS;
    function GetShellFolder(aShellFolder : TShellFolder): String;
    function GetShellFolderByName(aShellFolderName: String): String;
    function GetTempFile(Prefix: String): String;
    function GetTempFileAt(Dir, Prefix: String): String;
    function GetTempFolder(Prefix: String): String;
    function GetTempFolderAt(Dir, Prefix: String): String;
  Protected
  Public
    Constructor Create; Virtual;
    Function  WinExec32At(FileName : String; aPath : String; Visibility : Integer) : DWord; Virtual;
    Function  WinExecAndWait32(FileName : String; Visibility : Integer) : DWord; Virtual; Deprecated;
    Function  WinExecAndWait32At(FileName : String; aPath : String; Visibility : Integer) : DWord; Virtual; Deprecated;
    Function  WinExecEnv32(FileName : String; Env : String; Visibility : Integer) : DWord; Virtual;

    Function ShellDelete(Filename : String) : Boolean; Overload;
    Function ShellDelete(Files : TStrings) : Boolean; Overload;
  Public
    Property NetworkUserName                             : String Read fNetworkUserName;
    Property UserName                                    : String Read fUserName;
    Property Workstation                                 : String Read fWorkstation;
  Public
    Property ShellFolder[aShellFolder : TShellFolder]    : String Read GetShellFolder;
    Property ShellFolderByname[aShellFolderName: String]    : String Read GetShellFolderByName;
  Public
    Property TempFile[Prefix : String]                   : String Read GetTempFile;
    Property TempFileAt[Dir : String; Prefix : String]   : String Read GetTempFileAt;
    Property TempFolder[Prefix : String]                 : String Read GetTempFolder;
    Property TempFolderAt[Dir : String; Prefix : String] : String Read GetTempFolderAt;

  Public
    Property SystemPath                                  : String Read fSystemPath;
    Property WindowsPath                                 : String Read fWindowsPath;
    Property ProgramFilesPath                            : String Read fProgramFilesPath;
    Property TempPath                                    : String Read fTempPath;
    Property OS                                          : TidOS  Read fIDOS;
  End;

  TDDUDesktop = Class
  Private
    function GetBottom : Integer;
    function GetHeight : Integer;
    function GetLeft : Integer;
    function GetPath : String;
    function GetRight : Integer;
    function GetTop : Integer;
    function GetWidth : Integer;
  Protected
  Public
    Property Bottom : Integer Read GetBottom;
    Property Height : Integer Read GetHeight;
    Property Left   : Integer Read GetLeft;
    Property Path   : String  Read GetPath;
    Property Right  : Integer Read GetRight;
    Property Top    : Integer Read GetTop;    
    Property Width  : Integer Read GetWidth;  
  End;

  TDDUEnvironment= Class
  Private
    fEnvironment : TStringList;
    function GetText: String;
    function GetValue(Name: String): String;
    function GetValueDef(Name, aDefault: String): String;
    function GetExists(Name: String): Boolean;
    procedure SetValue(Name: String; const Value: String);
  Public
    Constructor Create; Virtual;
    Destructor Destroy; Override;

    Procedure LoadFromSystem;
    Procedure LoadFromStrings(Source : TStrings);

    Function Expand(Source : String) : String;
  Public
    Property Exists[Name : String]                      : Boolean Read GetExists;
    Property Value[Name : String]                       : String   Read GetValue Write SetValue;
    Property ValueDef[Name : String; aDefault : String] : String   Read GetValueDef;
    Property Text                                       : String   Read GetText;
  End;

  TDDUPath=Class
  Private
  Protected
  Public
    Function  AddBackslash(Const aPath : String) : String;
    Procedure CopyTree(aSourcePath : String; aDestPath : String);
    Function  CreatePath(Const APath : String) : Boolean;
    Function  DirExists(aPath : string) : Boolean;
    Function  FullPathToRelativePath(BasePath : String; FullPath : String) : String;
    Function  IsRoot(aPath : String) : Boolean;
    Procedure KillTree(aPath : String);
    Function  LoseBackslash(Const aPath : String) : String;
    Procedure MoveTree(aSourcePath : String; aDestPath : String);
    Function  ParentDir(Const aPath : String) : String;
    Function  RelativePathToFullPath(BasePath : String; RelativePath : String) : String;
    Function  SafeAddBackslash(Const aPath : String) : String;
    Procedure WalkTree(aPath : String; aMask : String; Recurse : Boolean; OnFile,OnFolder : TOnWalkTreeEvent);
  End;

  TDDURTTI = Class
  Private
  Protected
    Function EmitChar(What : Char) : String;
    Function Emit(Low,High : Char) : String;
    Function ReverseEmit(What : String) : Char;
  Public
    Function  CharSetToString(C : TCharSet) : String;
    Procedure CreateSet(Source : String; Var ResultSet : TCharSet);
    Function  EnumToString(Const Value; Info : PTypeInfo) : String;
    Function  SetToString(Const SetValue; Info : PTypeInfo) : String;
    Procedure StringToCharSet(S : String; Var C : TCharSet);
    Function  StringToEnum(Value : String; Info : PTypeInfo) : Integer;
    Procedure StringToSet(S : String; Info : PTypeInfo; Var SetValue);
  End;

  TDDUIntScale= Class
  Private
    fRatioX      : Extended;
    fRatioY      : Extended;
    fScaledWorld : TRect;
    fWorld       : TRect;
  Protected
  Public
    Constructor Create(aWorld : TRect; aScaledWorld : TRect);
    Function Scale(R : TRect) : TRect;

    Property World       : TRect Read fWorld;
    Property ScaledWorld : TRect Read fScaledWorld;
  End;

  TDDUEScale= Class
  Private
    fRatioX      : Extended;
    fRatioY      : Extended;
    fScaledWorld : TERect;
    fWorld       : TERect;
  Protected
  Public
    Constructor Create(aWorld : TERect; aScaledWorld : TERect);
    Function Scale(R : TERect) : TERect;

    Property World       : TERect Read fWorld;
    Property ScaledWorld : TERect Read fScaledWorld;
  End;

  TDDUFolderEnumerator=Class
  private
    fSearch    : String;
    fSearchRec : TSearchRec;
    fAttr      : INteger;
    fStarted   : Boolean;
  public
    constructor Create(aSearch : String; aAttr :Integer);
    Destructor Destroy; Override;
    function GetCurrent: TSearchRec;
    function MoveNext: Boolean;
    property Current: TSearchRec read GetCurrent;
  end;

  TDDUFolder=Class(TObject)
  Private
    fFolder : String;
    fMask   : String;
    fAttr   : Integer;
  Protected
  Public
    constructor Create(aFolder, aMask: String; aAttr : Integer);
    function GetEnumerator: TDDUFolderEnumerator; Virtual;

    Property Folder : String  Read fFolder Write fFolder;
    Property Mask   : String  Read fMask   Write fMask;
    Property Attr   : Integer Read fAttr   Write fAttr;
  End;


// Parsing functions
Procedure BreakApart(Source : String; Delim : TCharSet; Dest : TStrings);
Procedure CSVToStrings(Source : String; CSV : TStrings);
Function  GetMarkedText(Var Source : String; Out MarkedText : String; StartMarker,StopMarker : String) : Boolean;
Function  GetNextTerm(Var S : String; Seperator : String) : String;
Function  GetNextWord(Var S : String) : String;

Function  GetNextInt(Var S: String; Default: Integer =0) : Integer;
Function  GetNextFloat(Var S: String; Default:Extended =0) : Extended;
Function  GetNextCurr(Var S: String; Default:Currency=0) : Currency;

Function  LastPos(Substr : string; S : string) : Integer;
Procedure NormalBreakApart(Source : String; Delim : TCharSet; Dest : TStrings);
Function  Normalize(Const AString : String) : String;
Procedure ParsePhone(InPhone : String; Var Areacode : String; Var Phone : String; Var Ext : String);

Function StartsWith(Const Source : String; Const Prefix: String) : Boolean;
Function EndsWidth(Const Source : String; Const Suffix: String) : Boolean;
Function TextAfter(Const Source : String; Const Prefix: String) : String;
Function TextBefore(Const Source : String; Const Suffix: String) : String;
Function TextInside(Const Source : String; COnst Start,Finish : String) : String;

// Parameter parsing
Function  GetParamByFlag(Flag : String; Default : String) : String;
Function  IsParamFlagSet(Flag : String) : Boolean;
Function  ParamFlagPos(Flag : String) : Integer;

// Filename Manipulation
Function  ModifyFileName(Filename : String; Prefix,Suffix,Extension : String) : String ; Overload;
Function  ModifyFileName(Filename : String; Prefix,Suffix : String) : String ; Overload;
Function ModifyFileName(Filename : String; Prefix: String) : String; Overload;

Function  CanLockFile(FileName : String) : Boolean;

Procedure CloneStream(Origin : TStream; Dest : TStream);
function  GMTToLocal(GMT : TDateTime) : TDateTime;

Function  NumStrCompare(S1,S2 : String; CaseInsensitive : Boolean = True; IgnoreSpaces : Boolean = False) : Integer;

Function  IsCurrency(S : String) : Boolean;
Function  IsExtended(S : String) : Boolean;
Function  IsInt(S : String) : Boolean;
Function  IsDigits(S : String) : Boolean;

function  LocalToGMT(Local : TDateTime) : TDateTime;
Function  SafeFreeAndNil(Var O) : Boolean;

Function  BoolToStr(Condition : Boolean) : String;
Function  DDUBoolToStr(Condition : Boolean) : String;

Function  TryStrToBool(Const S: String; Var Value : Boolean) : Boolean;
Function  StrToBool(Const S : String) : Boolean;
Function  StrToBoolDef(Const S : String; Default : Boolean) : Boolean;

Function  TryDDUStrToBool(Const S: String; Var Value : Boolean) : Boolean;
Function  DDUStrToBool(Const S : String) : Boolean;
Function  DDUStrToBoolDef(Const S : String; Default : Boolean) : Boolean;

Function  CopyFilesByMask(aSourceDirectory,aDestDirectory : String; aMask : String; FailIfExists : Boolean=False) : Boolean;
Procedure DeleteFilesByMask(aDirectory : String; aMask : String);
Function  MoveFilesByMask(aSourceDirectory,aDestDirectory : String; aMask : String) : Boolean;

Function  AddBackslash(Const aPath : String) : String;
Function  CharSetToString(C : TCharSet) : String;
Function  CreatePath (Const APath : String; Agressive : Boolean=True ) : Boolean;
Procedure CreateSet(Source : String; Var ResultSet : TCharSet);
Function  DDUFileExists(aFilename: String) : Boolean;
Function  DirExists(aPath : string) : Boolean;
Function  EnumToString(Value : Cardinal; Info : PTypeInfo) : String; Overload;
Function  EnumToString(Const Value; Info : PTypeInfo) : String; Overload;

Function  FullPathToRelativePath(BasePath : String; FullPath : String) : String;
Function  GetEnv : String;
Function  GetEnvValue(aName : String; aDefault : String) : String;
procedure GetFileVersion(FileName: string; var Major1, Major2,Minor1, Minor2: Integer);
Function  GetFileVersionItem(Filename : String; Item : String) : String;
Function  GetFileVersionText(FileName : String) : String;
Function  GetTopFolder(Path : String) : String;
Function  GetNetUserName : String;
function  GetOperatingSystem: TIDOS;
Function  GetUserName : String;
Function  GetWorkStation : String;
Function  IsRoot(aPath : String) : Boolean;
Function  KillTree(DirName : String; Agressive : Boolean=False ) : Boolean;
Function  LoseBackslash(Const aPath : String) : String;
Function  MakeTempFile(Dir : String; Prefix : String) : String;
Function  MakeTempFolder(Dir : String; Prefix : String) : String;

Function  MoveTree(aSourcePath : String; aDestPath : String; DeleteSource : Boolean) : Boolean;
Function  ParentDir(Const aPath : String) : String;
Function  RecodePath(Const Source : String) : String; Deprecated;
Function  RelativePathToFullPath(BasePath : String; RelativePath : String) : String;
Function  SafeAddBackslash(Const aPath : String) : String;
Function  SetToString(Const SetValue; Info : PTypeInfo) : String;
Procedure StringToCharSet(S : String; Var C : TCharSet);
Function  StringToEnum(Value : String; Info : PTypeInfo) : Integer;
Procedure StringToSet(S : String; Info : PTypeInfo; Var SetValue);

Function FileTimeToDateTime(FileTime : TFileTime) : TDateTime;
Function DateTimeToFileTime(DateTime: TDateTime) : TFileTime;

Procedure TouchFile(Filename : String;
                    CreationTime            : TFileTime;
                    LastAccessTime          : TFileTime;
                    LastWriteTime           : TFileTime);
Procedure TouchFileFromFile(Filename : String; SourceFile : String);
Function GetFileTimes(Filename : String; Var CreationTime : TFileTime; Var LastAccessTime : TFileTime; Var LastWriteTime : TFileTime) : Boolean; Overload;
Function GetFileTimes(Filename : String; Var CreationTime : TDateTime; Var LastAccessTime : TDateTime; Var LastWriteTime : TDateTime) : Boolean; Overload;

Function  WinExec32At(FileName : String; aPath : String; Visibility : Integer) : DWord;
Function  WinExecAndWait32(FileName : String; Visibility : Integer) : DWord; Deprecated;
Function  WinExecAndWait32At(FileName : String; aPath : String; Visibility : Integer) : DWord; Deprecated;
Function  WinExecEnv32(FileName : String; Env : String; Visibility : Integer) : DWord;

Procedure AspectRatio(DisplayW, DisplayH : Integer; SourceW,SourceH : Extended; Var FinalW,FinalH : Integer); Overload;
Procedure AspectRatio(DisplayW, DisplayH : Extended; SourceW,SourceH : Extended; Var FinalW,FinalH : Extended); Overload;
Procedure ScaleERect(SourceRect : TERect; DestRect : TERect; Input : TERect; Var ScaledResult : TERect);

Function ERectToRect(ERect : TERect) : TRect;
Function RectToERect(Rect : TRect) : TERect;
Function ERect(Left,Bottom,Right,Top : Extended) : TERect;

Function PrintableTime(Milliseconds : Cardinal) : String; Overload;
Function PrintableTimeEx(Milliseconds : Cardinal) : String; Overload;

Function PrintableTime(aTime: TDateTime; ShowMS : Boolean=True) : String; Overload;
Function PrintableTimeEx(aTime : TDateTime; ShowMS: Boolean=True) : String; Overload;

Function KB_Value(Value : UInt64) : String;

Function NewGUIDString : String;

Function FirstNotBlank(Items : Array Of String) : String;
procedure SetCurrentThreadName(const Name: string);

Procedure FastHex(B : Byte; Var C1,C2 : AnsiChar); Inline;
Function FastUnHex(C : AnsiChar) : Byte; Inline;
function RightPos(const substr, str: String): Integer;

function CharStringIntersection(S1, S2: String): String;
Function CharStringIsEqual(S1,S2 : String) : Boolean;
Function CharStringSimplify(S1 : String) : String;
function CharStringUnion(S1, S2: String): String;

Function FastXMLFormat(XMLString: String): String;

function DDUMatchesMask(const TestSample, Mask: string): Boolean;

Procedure GetStringPropertyList(aObj : TObject; Names : TStrings);
Procedure MySetStrProp(Instance: TObject; const PropName: string; const Value: string);

// Code for working with bitmasks
Function ShiftCount(Mask : UInt64) : UInt8;
Function ShiftAndMask(Const Value, Mask : UInt64) : Uint64;
Function UnShiftAndMask(Const Value, Mask : UInt64) : Uint64;
Procedure SetBitField(Var Storage; Bytes : Int8; Const Mask : UInt64; Const Value : UInt64); Overload;
Function GetBitField(Var Storage; Bytes : Int8; Const Mask : UInt64) : Int64; Overload;
Function ReplaceMarkup(Const Source :String; Const Prefix : String; Values : TStrings) : String;

Var
  InternalDebugProc   : TInternalDebugProc;
  InternalDebugMethod : TInternalDebugMethod;
  SilentDebug         : Boolean;
  DDUSystem           : TDDUSystem;
  DDUDesktop          : TDDUDesktop;
  DDUEnvironment      : TDDUEnvironment;
  DDUPath             : TDDUPath;
  DDURTTI             : TDDURTTI;
  NoDebug             : Boolean;
  ConsoleDebug        : Boolean;

Implementation

Const
  sInvalidBoolean     = '''%s'' is not a valid boolean value.';
  sInvalidISODateTime = '''%s'' is not a valid ISO DateTime value.';
  sInvalidISODate     = '''%s'' is not a valid ISO Date value.';
  sInvalidISOTime     = '''%s'' is not a valid ISO Time value.';

Function BoolToStr(Condition : Boolean) : String;

Begin
  Result := DDUBoolToStr(Condition);
End;

Function DDUBoolToStr(Condition : Boolean) : String;

Begin
  Try
    Result := BooleanText[Condition];
  Except
    On E:Exception Do
    Begin
      Raise Exception.CreateFmt('BoolToStr Debug Exception (%s) -> %s',[E.ClassName,E.Message]);
    End;
  End;
End;

Procedure BreakApart(Source : String; Delim : TCharSet; Dest : TStrings);

Var
  CharCount             : Integer;
  DelimCount            : Integer;

Begin
  Dest.Clear;
  If Delim=[] Then
  Begin
    Include(Delim,#32);
  End;
  While Length(Source)>0 Do
  Begin
    { Elimate leading deliminators }
    DelimCount := 0;
    While CharInSet(Source[DelimCount+1],Delim) Do
    Begin
      Inc(DelimCount);
    End;
    Delete(Source,1,DelimCount);
    If Length(Source)=0 Then        // Trailing deliminators do not constitue
    Begin                           // additional information.
      Break;
    End;
    CharCount := 0;
    Repeat
      Inc(CharCount);
    Until (CharCount>=Length(Source)) Or CharInSet(Source[CharCount+1],Delim);
    Dest.Add(Copy(Source,1,CharCount));
    Delete(Source,1,CharCount);
  End;
End;

Function CanLockFile(FileName : String) : Boolean;

Var
  hFile                 : THandle;

Begin
  SetLastError(0);
  hFile := CreateFile(PChar(FileName),
                      Generic_Read or Generic_Write,
                      0,
                      Nil,
                      OPEN_EXISTING,
                      0,
                      0);
  Result := Not (hFile=INVALID_HANDLE_VALUE);
  If Result Then
  Begin
    CloseHandle(hFile);
  End;
End;

Procedure CloneStream(Origin : TStream; Dest : TStream);

Begin
  If (Origin Is TMemoryStream) And (Dest Is TMemoryStream) Then
  Begin
    Dest.Size := Origin.Size;
    Move(TMemoryStream(Origin).Memory^,TMemoryStream(Dest).Memory^,Dest.Size);
    Origin.Position := Origin.Size;
    Dest.Position := Dest.Size;
  End
  Else
  Begin
    Dest.Size := Origin.Size;
    Dest.Position := 0;
    Origin.Position := 0;
    Dest.CopyFrom(Origin,Origin.Size);
  End;
End;


function CharStringIntersection(S1, S2: String): String;

Var
  Loop                    : Integer;

begin
  Result := '';

  For Loop := 1 To Length(S1) Do
  Begin
    If Pos(S1[Loop],S2)<>0 Then Result := Result+S1[Loop];
  End;

  For Loop := 1 TO Length(S2) Do
  Begin
    If Pos(S2[Loop],S1)=0 Then Result := Result+S2[Loop];
  End;
  Result := CharStringSimplify(REsult);
end;

Function CharStringIsEqual(S1,S2 : String) : Boolean;

Var                     
  Loop                    : Integer;

Begin
  Result := False;

  For Loop := 1 To Length(S1) Do
  Begin
    If Pos(S1[Loop],S2)=0 Then Exit;
  End;
  For Loop := 1 To Length(S2) Do
  Begin
    If Pos(S2[Loop],S1)=0 Then Exit;
  End;

  Result := True;
End;

Function CharStringSimplify(S1 : String) : String;

Var
  Loop                    : Integer;

Begin
  Result := '';
  For Loop := 1 TO Length(S1) Do
  Begin
    If Pos(S1[Loop],Result)=0 Then Result := Result+S1[Loop];
  End;
End;

function CharStringUnion(S1, S2: String): String;

Var
  Loop                    : Integer;

begin
  Result := CharStringSimplify(S1);
  For Loop := Length(Result) DownTo 1 Do
  Begin
    If Pos(Result[Loop],S2)=0 Then Delete(Result,Loop,1);
  End;
end;

Procedure CSVToStrings(Source : String; CSV : TStrings);

Var
  At : Integer;

Begin
  CSV.Clear;
  While (source<>'') And (Source[1]='"') Do
  Begin
    Delete(Source,1,1);
    At := Pos('"',Source);
    CSV.Add(Copy(source,1,At-1));
    Delete(Source,1,At+1);
  End;
End;

Function FirstNotBlank(Items : Array Of String) : String;

Var
  Loop                    : Integer;

Begin
  Result := '';
  For Loop := Low(Items) To HIgh(Items) Do
  Begin
    If Trim(Items[Loop])<>'' Then
    Begin
      Result := Items[Loop];
      Break;
    End; 
  End;
End;

procedure GetFileVersion(FileName: string; var Major1, Major2,Minor1, Minor2: Integer);

var
  Info                  : Pointer;
  InfoSize              : DWORD;
  FileInfo              : PVSFixedFileInfo;
  FileInfoSize          : DWORD;
  Tmp                   : DWORD;

begin
  InfoSize := GetFileVersionInfoSize(PChar(FileName), Tmp);
  if (InfoSize=0) then
  Begin
    raise Exception.Create('Can''t get file version information for '+ FileName);
  End;
  GetMem(Info, InfoSize);
  try
    GetFileVersionInfo(PChar(FileName), 0, InfoSize, Info);
    VerQueryValue(Info, '\', Pointer(FileInfo), FileInfoSize);
    Major1 := FileInfo.dwFileVersionMS shr 16;
    Major2 := FileInfo.dwFileVersionMS and $FFFF;
    Minor1 := FileInfo.dwFileVersionLS shr 16;
    Minor2 := FileInfo.dwFileVersionLS and $FFFF;
  finally
    FreeMem(Info, FileInfoSize);
  end;
end;

Function GetFileVersionItem(Filename : String; Item : String) : String;

 var
   Buf                     : Pointer;
   BufSize                 : Cardinal;
   Len                     : Cardinal;
   Value                   : Pointer;

begin
  Result := '';
  BufSize := GetFileVersionInfoSize(PChar(Filename),BufSize);
  If BufSize > 0 Then
  Begin
    GetMem(Buf,BufSize);
    Try
      GetFileVersionInfo(PChar(Filename),0,BufSize,Buf);

      If VerQueryValue(Buf,PChar('\StringFileInfo\040904E4\'+Item),Value,Len) Then
      Begin
        If (Value<>Nil) And (Len>0) Then
        Begin
          SetLength(Result,Len);
          Move(Value^,Result[1],Len*SizeOf(Char));
        End;
      End;
    Finally
      FreeMem(Buf,BufSize);
    End;
  End;
End;

Function GetFileVersionText(FileName : String) : String;

Var
  Major1                : Integer;
  Major2                : Integer;
  Minor1                : Integer;
  Minor2                : Integer;

Begin
  GetFileVersion(FileName,Major1, Major2,Minor1, Minor2);
  Result := Format('Version %d.%d.%d Build %d',[Major1, Major2,Minor1,Minor2]);
End;

Function  GetTopFolder(Path : String) : String;

Begin
  Result := ExtractFileName( LoseBackSlash(Path) );
End;

Function GetMarkedText(Var Source : String; Out MarkedText : String; StartMarker,StopMarker : String) : Boolean;

Var
  At                      : Integer;
  At2                     : Integer;

Begin
  At := Pos(UpperCase(StartMarker),UpperCase(Source));
  If At<>0 Then
  Begin
    MarkedText := Copy(Source,At+Length(StartMarker),Length(Source)- (At+Length(StartMarker)));
    At2 := Pos(UpperCase(StopMarker),UpperCase(MarkedText));
    If (At2<>0) Then
    Begin
      SetLength(MarkedText,At2-1);
    End;
    Delete(Source,At,Length(StartMarker)+Length(MarkedText)+Length(StopMarker));
    Result := True;
  End
  Else
  Begin
    Result := False;
    MarkedText := '';
  End;
End;

Function  GetNextInt(Var S: String; Default: Integer =0) : Integer;

Var                     
  L                     : Integer;

Begin
  L := 0;
  While (S[L+1]<=#32) Do Inc(L);
  If (S[L+1]='+') Or (S[L+1]='-') Then Inc(L);

  While (S[L+1]>='0') And (S[L+1]<='9') Do Inc(L);

  Result := StrToIntDef(Trim(Copy(S,1,L)),Default);
  Delete(S,1,L);
End;

Function  GetNextFloat(Var S: String; Default:Extended =0) : Extended;

Var
  L                     : Integer;

Begin
  L := 0;
  While (S[L+1]<=#32) Do Inc(L);
  If (S[L+1]='+') Or (S[L+1]='-') Then Inc(L);

  While (S[L+1]>='0') And (S[L+1]<='9') Do Inc(L);
  If (S[L+1]='.') Then
  Begin
    Inc(L);
    While (S[L+1]>='0') And (S[L+1]<='9') Do Inc(L);
  End;
  Result := StrToFloatDef(Trim(Copy(S,1,L)),Default);
  Delete(S,1,L);
End;

Function  GetNextCurr(Var S: String; Default:Currency=0) : Currency;

Var
  L                     : Integer;

Begin
  L := 0;
  While (S[L+1]<=#32) Do Inc(L);
  If (S[L+1]='+') Or (S[L+1]='-') Then Inc(L);

  While (S[L+1]>='0') And (S[L+1]<='9') Do Inc(L);
  If (S[L+1]='.') Then
  Begin
    Inc(L);
    While (S[L+1]>='0') And (S[L+1]<='9') Do Inc(L);
  End;
  Result := StrToFloatDef(Trim(Copy(S,1,L)),Default);
  Delete(S,1,L);
End;

Function  GetNextTerm(Var S : String; Seperator : String) : String;

Var
  AT : Integer;

Begin
  At := Pos(Seperator,S);
  If At=0 Then
  Begin
    Result := S;
    S      := '';
  End
  Else
  Begin
    Result := Copy(S,1,At-1);
    Delete(S,1,At+Length(Seperator)-1);
  End;
End;

Function  GetNextWord(Var S: String) : String;

Var
  At                      : Integer;
  Loop                    : Integer;

Begin
  S := Trim(S);
  At := 0;
  For Loop := 1 to Length(S) Do
  Begin
    If (S[Loop]<=' ') Then
    Begin
      At := Loop;
      Break;
    End;
  End;
  If At=0 Then
  Begin
    Result := S;
    S := '';
  End
  Else
  Begin
    Result := Copy(S,1,At-1);
    S := Trim(Copy(S,At+1,Length(S)-At));
  End;
End;

Function GetParamByFlag (Flag : String; Default : String) : String;

Var
  Index                   : Integer;
  At                      : Integer;
  Work                    : String;

Begin
  Index := ParamFlagPos(Flag);
  If (Index<>-1) Then
  Begin
    Work := ParamStr(Index);
    At := Pos('=',Work);
    If At<>0 Then
    Begin
      Result := Copy(Work,At+1,Length(Work)-At);
    End
    Else
    Begin
      If (Index+1<=ParamCount) Then
      Begin
        Result := ParamStr(Index+1);
      End
      Else
      Begin
        Result := Default;
      End;
    End;
  End
  Else
  Begin
    Result := Default;
  End;
End;

function GMTToLocal(GMT : TDateTime) : TDateTime;

Var
  FileDate              : Integer;
  GMTTime               : TFileTime;
  LocalTime             : TFileTime;

Begin
  FileDate := DateTimeToFileDate(GMT);
  DosDateTimeToFileTime(LongRec(FileDate).Hi, LongRec(FileDate).Lo, GMTTime);
  FileTimeToLocalFileTime(GMTTime,LocalTime);
  FileTimeToDosDateTime(LocalTime,LongRec(FileDate).Hi, LongRec(FileDate).Lo);
  Result := FileDateToDateTime(FileDate);
End;

Function ifop(Condition : Boolean; IfTrue : Integer; IfFalse : Integer) : Integer; Overload;

Begin
  If Condition Then Result := iftrue Else Result := ifFalse;
End;

Function ifop(Condition : Boolean; IfTrue : Int64; IfFalse : Int64) : Int64; Overload;

Begin
  If Condition Then Result := iftrue Else Result := ifFalse;
End;

Function IsCurrency(S : String) : Boolean;

Var
  I                       : Currency;

Begin
  Result := TryStrToCurr(S,I);
End;

Function IsExtended(S : String) : Boolean;

Var
  I                       : Extended;

Begin
  Result := TryStrToFloat(S,I);
End;

Function IsInt(S : String) : Boolean;

Var
  I                       : Integer;

Begin
  Result := TryStrToInt(S,I);
End;

Function  IsDigits(S : String) : Boolean;

Var
  Loop                    : Integer;

Begin
  Result := True;
  For Loop := 1 To Length(S) Do
  Begin
    If (S[Loop]<'0') OR (S[Loop]>'9') Then
    Begin
      Result := False;
      Break;
    End;
  End;
End;

Function IsParamFlagSet (Flag : String) : Boolean;

Var
  At                    : Integer;
  Work                  : String;
  prefix                : String;
  PostFix               : String;

Begin
  Result := False;
  At := ParamFlagPos(Flag);
  If (At<>-1) Then
  Begin
    Work := ParamStr(At);
    At := Pos('=',Work);
    If At<>0 Then
    Begin
      SetLength(Work,At-1);
    End;
    Prefix := Copy(Work,1,1);
    If InStringList(Prefix,['/','+','-']) Then
    Begin
      Delete(Work,1,1);
      PostFix := Copy(Work,Length(Work),1);
      If InStringList(PostFix,['+','-']) Then
      Begin
        Delete(Work,Length(Work),1);
      End
      Else
      Begin
        PostFix := '+';
      End;
      If (CompareText(Work,Flag)=0) Then
      Begin
        Result := (PostFix='+');
      End;
    End;
  End;
End;

Function LastPos(Substr: string; S: string): Integer;

Var
  Loop                    : Integer;
  SL                      : Integer;

Begin
  Result := 0;
  SL := Length(SubStr);
  For Loop := Length(S)-SL+1 Downto 1 Do
  Begin
    If Copy(S,Loop,SL)=SubStr Then
    Begin
      Result := Loop;
      Break;
    End;
  End;
End;

function LocalToGMT(Local : TDateTime) : TDateTime;

Var
  FileDate              : Integer;
  GMTTime               : TFileTime;
  LocalTime             : TFileTime;

Begin
  FileDate := DateTimeToFileDate(Local);
  DosDateTimeToFileTime(LongRec(FileDate).Hi, LongRec(FileDate).Lo, LocalTime);
  LocalFileTimeToFileTime(LocalTime,GMTTime);
  FileTimeToDosDateTime(GMTTime,LongRec(FileDate).Hi, LongRec(FileDate).Lo);
  Result := FileDateToDateTime(FileDate);
End;

Function ModifyFileName(Filename : String; Prefix,Suffix,Extension : String) : String;

Begin
  Result := ExtractFilePath(FileName)+Prefix+ChangeFileExt(ExtractFileName(FileName),'')+Suffix+Extension;
End;

Function ModifyFileName(Filename : String; Prefix,Suffix : String) : String;

Begin
  Result := ModifyFileName(FileName,Prefix,Suffix,ExtractFileExt(FileName));
End;

Function ModifyFileName(Filename : String; Prefix: String) : String;

Begin
  Result := ModifyFileName(FileName,Prefix,ExtractFileExt(FileName));
End;

Procedure NormalBreakApart(Source : String; Delim : TCharSet; Dest : TStrings);

Var
  CharCount             : Integer;        // Count of valid characters
  Trailing              : Boolean;

Begin
  Dest.Clear;
  If Delim=[] Then
  Begin
    Include(Delim,#32);
  End;
  While Length(Source)>0 Do
  Begin
    CharCount := 1;
    While (CharCount<Length(Source)) And (Not CharInSet(Source[CharCount],Delim)) Do
    Begin
      Inc(CharCount);
    End;
    If  CharInSet(Source[CharCount],Delim) Then
    Begin
      Dest.Add(Copy(Source,1,CharCount-1));
      Trailing := True;
    End
    Else
    Begin
      Dest.Add(Copy(Source,1,CharCount));
      Trailing := False;
    End;
    Delete(Source,1,CharCount);
    If (Source='') And Trailing Then
    Begin
      Dest.Add('');
      Source := '';
    End;
  End;
End;

Function Normalize(Const AString : String) : String;
//
// Purpose : Trim, Single Space, and convert to upper case, also removes double punctation.
//
Var
  Loop                  : Integer;
  Ch                    : Char;

Begin
  Result := UpperCase(Trim(AString));
  // Elimate double punctuation
  For Loop := Length(Result) Downto 2 Do
  Begin
    Ch := Result[Loop];
    If CharInSet(Ch,TPunctuation) And (Result[Loop-1]=Ch) Then
    Begin
      Delete(Result,Loop,1);
    End;
  End;
  // Elimate spaces before punctuation.
  For Loop := Length(Result)-1 Downto 1 Do
  Begin
    Ch := Result[Loop];
    If (Ch=#32) And CharInSet(Result[Loop+1],TPunctuation) Then
    Begin
      Delete(Result,Loop,1);
    End;
  End;
  // Elimate spaces after punctuation.
  For Loop := Length(Result)-1 Downto 1 Do
  Begin
    Ch := Result[Loop];
    If CharInSet(Ch,TPunctuation) And (Ch<>#32) And (Result[Loop+1]=#32) Then
    Begin
      Delete(Result,Loop+1,1);
    End;
  End;
End;

// Written by Clinton R. Johnson (xepol@xepol.com) and Mark Spankus (mark@cs.wisc.edu)
Function  NumStrCompare(S1,S2 : String; CaseInsensitive : Boolean=True;  IgnoreSpaces : Boolean=False) : Integer;

Var
  S1At                : PChar;
  S2At                : PChar;
  S1V                 : LongInt;
  S2V                 : LongInt;

Begin
  S1At := PChar(S1);
  S2At := PChar(S2);
  Result := 0;

  If IgnoreSpaces Then
  Begin
    While (S1At^=#32) Do Inc(S1At);
    While (S2At^=#32) Do Inc(S2At);
  End;

  While (Result=0) And (S1At^<>#0) And (S2At^<>#0) Do
  Begin
    If (S1At^<='9') And (S2At^<='9') And
       (S1At^>='0') And (S2At^>='0') Then
    Begin
      //  Need to search forward.
      S1V := 0;
      While (S1At^<='9') And (S1At^>='0') Do
      Begin
        S1V := (S1V*10)+Byte(S1At^)-Byte('0');
        Inc(S1At);
      End;
      S2V := 0;
      While (S2At^<='9') And (S2At^>='0') Do
      Begin
        S2V := (S2V*10)+Byte(S2At^)-Byte('0');
        Inc(S2At);
      End;
      Result := S1V-S2V;
    End
    Else
    Begin
      If CaseInsensitive Then
      Begin
        Result := Byte(Upcase(S1At^))-Byte(Upcase(S2At^));
      End
      Else
      Begin
        Result := Byte(S1At^)-Byte(S2At^);
      End;
      Inc(S1At);
      Inc(S2At);
    End;

    If IgnoreSpaces Then
    Begin
      While (S1At^=#32) Do Inc(S1At);
      While (S2At^=#32) Do Inc(S2At);
    End;
  End;

  If (Result=0) Then
  Begin
    Result := Byte(S1At^)-Byte(S2At^);
  End;

  If (Result<0) Then
  Begin
    Result := -1;
  End
  Else If (Result>1) Then
  Begin
    Result := 1;
  End;
End;
//
// C Function equivanlent :
//
// int NumStCompare(LPCSTR str1, LPCSTR str2, BOOL IgnoreCase, Bool IgnoreSpace)
// {
//   LPCSTR S1=str1;
//   LPCSTR S2=str2;
//   int S1V;
//   int S2V;
//   int Result=0;
//
//   if (!str1 || !str2) return 0; // No null pointers!
//
//   if (IgnoreSpaces)
//   {
//     while (isspace(*S1)) S1++;
//     while (isspace(*S2)) S2++;
//   }
//
//   while (!Result && *S1 && *S2)
//   {
//     if ((*S1<='9') && (*S2<='9') && (*S1>='0') && (*S2>='0'))
//     {
//       S1V=S2V=0;
//       while ((*S1<='9') && (*S1>='0')) S1V=(S1V*10)+*S1++-'0';
//       while ((*S2<='9') && (*S2>='0')) S2V=(S2V*10)+*S2++-'0';
//       Result=S1V-S2V;
//     } else {
//       Result= (IgnoreCase)?upcase(*S1++):*S1++ - (IgnoreCase)?upcase(*S2++):*S2++;
//     }
//     if (IgnoreSpaces)
//     {
//       while (isspace(*S1)) S1++;
//       while (isspace(*S2)) S2++;
//     }
//  }
//  if (!Result) Result=(int)*S1-(int)*S2;
//  if (Result<0) Result=-1;
//  if (Result>0) Result=1;
//  return Result;
//}

Function ParamFlagPos(Flag : String) : Integer;

Var
  Loop                  : Integer;
  Work                  : String;
  prefix                : String;
  PostFix               : String;
  At                    : Integer;

Begin
  Result := -1;
  For Loop := 1 To ParamCount Do
  Begin
    Work := ParamStr(loop);
    
    At := Pos('=',Work);
    If At<>0 Then
    Begin
      SetLength(Work,At-1);
    End;

    Prefix := Copy(Work,1,1);
    If InStringList(Prefix,['/','+','-']) Then
    Begin
      Delete(Work,1,1);
      PostFix := Copy(Work,Length(Work),1);
      If InStringList(PostFix,['+','-']) Then
      Begin
        Delete(Work,Length(Work),1);
      End;
      If (CompareText(Work,Flag)=0) Then
      Begin
        Result := Loop;
        Break;
      End;
    End;
  End;
End;

Procedure ParsePhone(InPhone : String; Var Areacode : String; Var Phone : String; Var Ext : String);

Var
 TDigitSet             : Set Of AnsiChar;
 TPhoneSet             : Set Of AnsiChar;
 TermSet               : Set Of AnsiChar;

Begin
  TDigitSet             := ['0'..'9'];
  TPhoneSet             := [#32,'-','.','0'..'9'];
  TermSet               := [']','}',')'];

  AreaCode := '';
  Phone := '';
  Ext := '';
  InPhone := Trim(Uppercase(InPhone));
  If Length(InPhone)=0 Then
    Exit;

  If  (InPhone[1]='1') Or (Copy(InPhone,1,2)='00') Or (Copy(InPhone,1,2)='01') Or
      (Not CharInSet(InPhone[1],TDigitSet)) Or
      (Pos('(',InPhone)=1) Or (Pos('+',InPhone)=1) Or
      ((Pos(' ',InPhone)<>0) And ( (Pos(' ',InPhone)<Pos('-',InPhone)) OR (Pos(' ',InPhone)<Pos('.',InPhone)) )   ) Then
  Begin
    If  (Copy(InPhone,2,1)='+') Then
    Begin
      Include (TDigitSet,'-');  // For international dialing
      Include (TDigitSet,'+');  // For international dialing
    End;
    If (Copy(InPhone,1,1)='+') Then
    Begin
      AreaCode := '+';
      Include (TDigitSet,'+');  // For international dialing
    End;

    If (InPhone[1]='1') Then
    Begin
      If (Copy(InPhone,1,2)='1-') Or (Copy(InPhone,1,2)='1.') Then
      Begin
        Delete (InPhone,1,2);
      End
      Else
      Begin
        AreaCode := '1';
        Delete (InPhone,1,1);
      End;
    End Else If (Copy(InPhone,1,2)='00') Or (Copy(InPhone,1,2)='01') Then
    Begin
      AreaCode := Copy(InPhone,1,2);
      Delete(InPhone,1,2);
    End;

    While (Length(InPhone)<>0) And (Not CharInSet(InPhone[1],TDigitSet)) Do
    Begin
      Delete(InPhone,1,1);
    End;

    While (Length(InPhone)<>0) And CharInSet(InPhone[1],TDigitSet) Do
    Begin
      AreaCode := AreaCode+Copy(InPhone,1,1);
      Delete(InPhone,1,1);
    End;

    While (Length(InPhone)<>0) And (Not CharInSet(InPhone[1],TDigitSet)) Do
    Begin
      Delete(InPhone,1,1);
    End;
  End;
  Exclude(TDigitSet,'-'); // For international dialing
  Exclude(TDigitSet,'+'); // For international dialing

  InPhone := Trim(InPhone);
  While (Length(InPhone)<>0) And CharInSet(InPhone[1],TPhoneSet) Do
  Begin
    If (InPhone[1]='-') OR (InPhone[1]='.') Then
    Begin
      Exclude(TPhoneSet,#32);
    End;
    If CharInSet(InPhone[1],TDigitSet) Then
    Begin
      Phone := Phone+Copy(InPhone,1,1);
    End;

    Delete(InPhone,1,1);
  End;
  Ext := Trim(InPhone);
  While (Length(Ext)<>0) And (Not CharInSet(Ext[1],TDigitSet)) Do
  Begin
    Delete(Ext,1,1);
  End;
  If (Length(Ext)<>0) And CharInSet(Ext[Length(Ext)],TermSet) Then
  Begin
    Delete(Ext,Length(Ext),1);
  End;
End;

Function StartsWith(Const Source : String; Const Prefix: String) : Boolean;

Begin
  Result := SameText( Copy(Source,1,Length(Prefix)),Prefix);
End;

Function EndsWidth(Const Source : String; Const Suffix: String) : Boolean;

Begin
  Result := SameText( Copy(Source,Length(Source)-Length(Suffix)+1,Length(Suffix)) , Suffix );
End;

Function TextAfter(Const Source : String; Const Prefix: String) : String;

Var
  At                      : Integer;

Begin
  If StartsWith(Source,Prefix) Then
  Begin
    Result := Copy(Source,Length(Prefix)+1,Length(Source)-Length(Prefix));
  End
  Else
  Begin
    At := Pos(Prefix,Source);
    If At=0 Then
    Begin
      Result := '';
    End
    Else
    begin
      Result := Copy(Source,At+Length(Prefix),Length(Source)-(At+Length(Prefix)-1));
    end;
  End;
End;

Function TextBefore(Const Source : String; Const Suffix: String) : String;

Var                     
  At                      : Integer; 

Begin
  If EndsWidth(Source,Suffix) Then
  Begin
    Result := Copy(Source,1,Length(Source)-Length(Suffix));
  End
  Else
  Begin
    At := Pos(Suffix,Source);
    If At=0 Then
    Begin
      Result := '';
    End
    Else
    begin
      Result := Copy(Source,1,At-1);
    end;
  End;
End;

Function TextInside(Const Source : String; COnst Start,Finish : String) : String;

Var
  At1,At2 : Integer;

Begin
  At1 := Pos(Start,Source);
  At2 := PosEx(Finish,Source,At1+Length(Start));
  If (At1<>0) And (At2<>0) Then
  Begin
//  12345678
//  abc[123]
//
//  At1=4
//  At2=8
//  Len= At2-At1-Len(Start)  8-4-1=3
    Result := Copy(Source,At1+Length(Start),At2-At1-Length(Start));
  End
  Else
  Begin
    Result := '';
  End;
End;

Function  SafeFreeAndNil(Var O) : Boolean;

Begin
  Try
    FreeAndNil(O);
    Result := True;
  Except
    Result := False;
  End;
End;

Function  TryStrToBool(Const S: String; Var Value : Boolean) : Boolean;

Begin
  Result := TryDDUStrToBool(S,Value);
End;

Function StrToBool(Const S : String) : Boolean;

Begin
  Result := DDUStrToBool(S);
End;

Function StrToBoolDef(Const S : String; Default : Boolean) : Boolean;

Begin
  Result := DDUStrToBoolDef(S,Default);
End;

Function  TryDDUStrToBool(Const S: String; Var Value : Boolean) : Boolean;

Begin
  If (InStringList(UpperCase(S),['1','True','T','Y','Yes'])) Then
  Begin
    Value := True;
    Result := True
  End Else If (InStringList(UpperCase(S),['0','False','F','N','No'])) Then
  Begin
    Value := False;
    Result := True;
  End Else
  Begin
    Value := False;
    Result := False;
  End;
End;

Function DDUStrToBool(Const S : String) : Boolean;

Begin
  If Not TryDDUStrToBool(S,Result) Then
  Begin
    Raise EConvertError.Createfmt(sInvalidBoolean,[S]);
  End;
End;

Function DDUStrToBoolDef(Const S : String; Default : Boolean) : Boolean;

Begin
  If Not TryDDUStrToBool(S,Result) Then
  Begin
    Result := Default;
  End;
End;

{ Deprecated Functions }

Function AddBackslash(Const aPath :String) : String;

Begin
  Result := DDUPath.AddBackslash(aPath);
End;

Function CharSetToString(C : TCharSet) : String;

Begin
  Result := DDURTTI.CharSetToString(C);
End;

Function CopyFilesByMask(aSourceDirectory,aDestDirectory : String; aMask : String; FailIfExists : Boolean=False) : Boolean;

Var
  SearchRec               : TSearchRec;
  DosError                : Integer;

Begin
  Result := True;
  DosError := FindFirst(aSourceDirectory+aMask,faAnyFile,SearchRec);
  While (DosError=0) Do
  Begin
    Try
      If Not CopyFile(PChar(aSourceDirectory+SearchRec.Name),PChar(ADestDirectory+SearchRec.Name),FailIfExists) Then
      Begin
        Result := False;
      End;
    Except
      Result := False
    End;
    DosError := FindNExt(SearchRec);
  End;
  FindClose(SearchRec);
End;
Function CreatePath (Const APath : String; Agressive : Boolean ) : Boolean;

Var
  Count                   : Integer;

Begin
  If (aPath<>'') And (Not DirectoryExists(AddBackslash(aPath) )) Then
  Begin
    Count := 0;
    Repeat
      If (Count<>0) Then
      Begin
//        Debug('Aggresive Create path, pass %d - %s',[Count+1,aPath]);
        Sleep(100);
      End;
      Result := DDUPath.CreatePath(aPath);
      Inc(Count);
    Until Result Or ((Not Agressive) Or (Count>=20)) ;
  End
  Else
  Begin
    Result := True;
  End;
End;

Procedure CreateSet(Source : String; Var ResultSet : TCharSet);

Begin
  DDURTTI.CreateSet(Source,ResultSet);
End;

Function DDUFileExists(aFilename: String) : Boolean;

Var
  SearchRec : TSearchRec;
  DosError  : Integer;

Begin
  If (Pos('*',aFilename)<>0) Or (Pos('?',aFilename)<>0) Then
  Begin
    Result := False;
    DosError := FindFirst(aFilename,faAnyFile,SearchRec);
    While (Not Result) And (DosError=0) Do
    Begin
      Result := ((SearchRec.Attr And faDirectory)=0);
      DosError := FindNext(SearchRec);
    End;
    FindClose(SearchRec);
  End
  Else
  Begin
    Result := FileExists(aFilename);
  End;
End;

Procedure DeleteFilesByMask(aDirectory : String; aMask : String);

Var
  SearchRec               : TSearchRec;
  DosError                : Integer;

Begin
  DosError := FindFirst(aDirectory+aMask,faAnyFile,SearchRec);
  While (DosError=0) Do
  Begin
    Try
      DeleteFile(aDirectory+SearchRec.Name);
    Except
    End;

    DosError := FindNExt(SearchRec);
  End;
  FindClose(SearchRec);
End;

Function DirExists(aPath : string): Boolean;

Begin
  Result := DDUPath.DirExists(aPath);
end;

Function EnumToString(Value : Cardinal; Info : PTypeInfo) : String;

Begin
  Result := DDURTTI.EnumToString(Value,Info);
End;

Function  EnumToString(Const Value; Info : PTypeInfo) : String;

Var
  Data                    : PTypeData;
  C                       : Cardinal;

Begin
  Data := GetTypeData(Info);

  Case Info^.Kind Of
    tkEnumeration : Begin
                      If Data.MaxValue<=$ff Then
                      Begin
                        C:= Byte(Value);
                      End Else If Data.MaxValue<=$ffff Then
                      Begin
                        C := Word(Value);
                      End Else //If Data.MaxValue<=$ffffffff Then
                      Begin
                        C:=  Cardinal(Value);
                      End;
                      Result := DDURTTI.EnumToString( C,Info);

                    End;
    tkInteger     : Result := DDURTTI.EnumToString( Integer(Value),Info);
  End;
End;

Function FullPathToRelativePath(BasePath : String; FullPath: String) : String;

Begin
  Result := DDUPath.FullPathToRelativePath(BasePath,FullPath);
End;

Function GetEnv : String;

begin
  Result := DDUEnvironment.Text;
End;

Function  GetEnvValue(aName : String; aDefault : String) : String;

Begin
  Result := DDUEnvironment.GetValueDef(aName,aDefault);
End;

Function GetNetUserName : String;

Begin
  Result:= DDUSystem.NetworkUserName;
end;

function GetOperatingSystem: TIDOS;
Begin
  Result := DDUSystem.OS;
end;


Function GetUserName : String;

begin
  Result := DDUSystem.UserName;
end;

Function  GetWorkStation : String;

begin
  Result := DDUSystem.Workstation;
end;

Function  IsRoot(aPath : String) : Boolean;

Begin
  Result := DDUPath.IsRoot(aPath);
End;

Function KillTree(DirName : String; Agressive : Boolean=False ) : Boolean;

Var
  Count                   : Integer;

Begin
  Count := 0;
  Repeat
    If (Count<>0) Then
    Begin
      Sleep(100);
    End;

    DDUPath.KillTree(DirName);

    Result := Not DirExists(DirName);
    Inc(Count);
  Until Result Or ((Not Agressive) Or (Count>=20)) ;
End;

Function LoseBackslash(Const aPath :String) : String;

Begin
  Result := DDUPath.LoseBackslash(aPath);
End;

Function MakeTempFile(Dir : String; Prefix : String) : String;

Begin
  Result := DDUSystem.TempFileAt[Dir,Prefix];
End;

Function  MakeTempFolder(Dir : String; Prefix : String) : String;

Begin
  Result := DDUSystem.TempFolderAt[Dir,Prefix];
End;

Function MoveFilesByMask(aSourceDirectory,aDestDirectory : String; aMask : String) : Boolean;

Var
  SearchRec               : TSearchRec;
  DosError                : Integer;

Begin
  Result := True;
  DosError := FindFirst(aSourceDirectory+aMask,faAnyFile,SearchRec);
  While (DosError=0) Do
  Begin
    Try
      DeleteFile(ADestDirectory+SearchRec.Name);
      If Not MoveFile(PChar(aSourceDirectory+SearchRec.Name),PChar(ADestDirectory+SearchRec.Name)) Then
      Begin
        Result := False;
      End;
    Except
      Result := False
    End;
    DosError := FindNext(SearchRec);
  End;
  FindClose(SearchRec);
End;


Function MoveTree(aSourcePath : String; aDestPath : String; DeleteSource : Boolean) : Boolean;

Begin
  Try
    If DeleteSource Then
    Begin
      DDUPath.MoveTree(aSourcePath,aDestPath);
    End
    Else
    Begin
      DDUPath.CopyTree(aSourcePath,aDestPath);
    End;
    Result := True;
  Except
    Result := False;
  End;
End;

Function ParentDir(Const aPath : String) : String;

Begin
  Result := DDUPath.ParentDir(aPath);
End;

Function RecodePath(Const Source : String) : String;

Var
  Marker                : String;

Begin
  Result := Source;
  If (Copy(Result,1,1)='<') And (Pos('>',Result)<>0) Then
  Begin
    Delete(Result,1,1);
    Marker := UpperCase(Result);
    SetLength(Marker,Pos('>',Marker)-1);
    Delete(Result,1,Length(Marker)+1);
    If Copy(Result,1,1)='\' Then
    Begin
      Delete(Result,1,1);
    End;
    Result := DDUSystem.ShellFolderByName[Marker]+Result;
  End;
End;

Function RelativePathToFullPath(BasePath : String; RelativePath : String) : String;

Begin
  Result := DDUPath.RelativePathToFullPath(BasePath,RelativePath);
end;

Function SafeAddBackslash(Const aPath :String) : String;

Begin
  Result := DDUPath.SafeAddBackslash(aPath);
End;

Function SetToString(Const SetValue; Info : PTypeInfo) : String;

Begin
  Result := DDURTTI.SetToString(SetValue,Info);
End;

Procedure StringToCharSet(S : String; Var C : TCharSet);

Begin
  DDURTTI.StringToCharSet(S,C);
End;

Function  StringToEnum(Value : String; Info : PTypeInfo) : Integer;

Begin
  Result := DDURTTI.StringToEnum(Value,Info);
End;

Procedure StringToSet(S : String; Info : PTypeInfo; Var SetValue);

Begin
  DDURTTI.StringToSet(S,Info,SetValue);
End;

function FileTimeToDateTime(FileTime: TFileTime): TDateTime;

var
  LocalFileTime           : TFileTime;
  SystemTime              : TSystemTime;

begin
  FileTimeToLocalFileTime(FileTime, LocalFileTime);
  FileTimeToSystemTime(LocalFileTime, SystemTime);
  Result := SystemTimeToDateTime(SystemTime) ;
end;

Function DateTimeToFileTime(DateTime: TDateTime) : TFileTime;

var
  LocalFileTime           : TFileTime;
  SystemTime              : TSystemTime;

Begin
  DateTimeToSystemTime(DateTime,SystemTime);
  SystemTimeToFileTime(SystemTime,LocalFileTime);
  LocalFileTimeToFileTime(LocalFileTime,Result);
End;

Procedure TouchFile(Filename : String;
                    CreationTime            : TFileTime;
                    LastAccessTime          : TFileTime;
                    LastWriteTime           : TFileTime);


Var
  aFile                   : HFile;

Begin
  aFile := CreateFile(PChar(Filename),GENERIC_WRITE,0,Nil,OPEN_EXISTING,0,0);
  If aFile<>INVALID_HANDLE_VALUE THen
  Begin
    SetFileTime(aFile,@CreationTime,@LastAccessTime,@LastWriteTIme);
    CloseHandle(aFile);
  End;
End;

Procedure TouchFileFromFile(Filename : String; SourceFile : String);

Var
  aFile                   : HFile;
  CreationTime            : TFileTime;
  LastAccessTime          : TFileTime;
  LastWriteTime           : TFileTime;

Begin
  aFile := CreateFile(PChar(SourceFile),GENERIC_READ ,0,Nil,OPEN_EXISTING,0,0);
  If (aFile<>INVALID_HANDLE_VALUE) Then
  Begin
    GetFileTime(aFile,@CreationTime,@LastAccessTime,@LastWriteTIme);
    CloseHandle(aFile);
    TouchFile(Filename,CreationTime,LastAccessTime,LastWriteTIme);
  End;
End;

Function GetFileTimes(Filename : String; Var  CreationTime            : TFileTime;
                                         Var  LastAccessTime          : TFileTime;
                                         Var  LastWriteTime           : TFileTime) : Boolean;

Var
  aFile                   : HFile;

Begin
  aFile := CreateFile(PChar(Filename),GENERIC_READ ,0,Nil,OPEN_EXISTING,0,0);
  If (aFile<>INVALID_HANDLE_VALUE) Then
  Begin
    GetFileTime(aFile,@CreationTime,@LastAccessTime,@LastWriteTIme);
    CloseHandle(aFile);
    Result := True;
  End
  Else
  Begin
    Result := False;
  End;
End;

Function GetFileTimes(Filename : String; Var CreationTime : TDateTime; Var LastAccessTime : TDateTime; Var LastWriteTime : TDateTime) : Boolean;

Var
  F1,F2,F3 : TFileTime;

Begin
  Result := GetFileTimes(Filename,F1,F2,F3);
  If Result Then
  Begin
    CreationTime   := FileTimeToDateTime(F1);
    LastAccessTime := FileTimeToDateTime(F2);
    LastWriteTime  := FileTimeToDateTime(F3);
  End;
End;

Function  WinExec32At(FileName : String; aPath : String; Visibility : Integer) : DWord;

Begin
  Result := DDUSystem.WinExec32At(Filename,aPath,Visibility);
End;

Function  WinExecAndWait32(FileName : String; Visibility : Integer) : DWord;

Begin
  Result := DDUSystem.WinExecAndWait32(Filename,Visibility);
End;

Function  WinExecAndWait32At(FileName : String; aPath : String; Visibility : Integer) : DWord;

Begin
  Result := DDUSystem.WinExecAndWait32At(Filename,aPath,Visibility);
End;

Function  WinExecEnv32(FileName : String; Env : String; Visibility : Integer) : DWord;

Begin
  Result := DDUSystem.WinExecEnv32(Filename,Env,Visibility);
End;

{ TDDUSystem }

constructor TDDUSystem.Create;

Var                     
  Size                    : Cardinal;  
  R                       : TRegistry; 

begin
{ TODO -cUnicode : Check if this works with unicode. }
  fIDOS := GetOperatingSystem;

  Size := 255;
  SetLength(fUserName,Size);
  If WinAPI.Windows.GetUserName(PChar(fUserName),Size) Then
  Begin
    fUserName := PChar(fUserName);
  End
  Else
  Begin
    fUserName := '';
  End;

  Size := 255;
  SetLength(fWorkStation,Size);
  If WinAPI.Windows.GetComputerName(PChar(fWorkStation),Size) Then
  Begin
    fWorkStation := PChar(fWorkStation);
  End
  Else
  Begin
    fWorkStation := '';
  End;
  fNetworkUserName := fUserName+'@'+fWorkstation;

  SetLength(fSystemPath,MAX_PATH);
  GetSystemDirectory(PChar(fSystemPath),MAX_PATH);
  SetLength(fSystemPath,StrLen(PChar(fSystemPath)));

  SetLength(fTempPath,MAX_PATH);
  GetTempPath(MAX_PATH,PChar(fTempPath));
  SetLength(fTempPath,StrLen(PChar(fTempPath)));

  SetLength(fWindowsPath,MAX_PATH);
  GetWindowsDirectory(PChar(fWindowsPath),MAX_PATH);
  SetLength(fWindowsPath,StrLen(PChar(fWindowsPath)));


  R := TRegistry.Create;
  Try
    R.RootKey := HKEY_LOCAL_MACHINE;
    R.Access := KEY_QUERY_VALUE;
    If R.OpenKey('\Software\Microsoft\Windows\CurrentVersion',False) Then
    Begin
      fProgramFilesPath := SafeAddBackslash(R.ReadString('ProgramFilesDir') ); 
      R.CloseKey;
    End;
  Finally
    FreeAndNil(R);
  end;
  If fProgramFilesPath='' Then
  Begin
    fProgramFilesPath := 'C:\Program Files\';
  End;
End;

function TDDUSystem.GetOperatingSystem: TIDOS;

var
  osVerInfo               : TOSVersionInfo;
  majorVer                : Integer;
  minorVer                : Integer;
//  ComplexVer              : Currency;

begin
  FillChar(osVerInfo,SizeOf(osVerInfo),0);
  osVerInfo.dwOSVersionInfoSize := SizeOf(TOSVersionInfo);
  if GetVersionEx(osVerInfo) then
  begin
    majorVer := osVerInfo.dwMajorVersion;
    minorVer := osVerInfo.dwMinorVersion;
    case osVerInfo.dwPlatformId of
      VER_PLATFORM_WIN32_NT      : begin { Windows NT/2000 }
                                     if majorVer <= 4 then
                                      Result := idOSWinNT
                                     else if (majorVer = 5) and (minorVer = 0) then
                                       Result := idOSWin2000
                                     else if (majorVer = 5) and (minorVer = 1) then
                                       Result := idOSXP
                                     else if (majorVer = 6) And (minorVer=0) Then
                                       Result := idOSVista
                                     else if (majorVer = 6) And (minorVer=1) Then
                                       Result := idOSWin7
                                     else if (majorVer = 6) And (minorVer=2) Then
                                       Result := idOSWin8
                                     else if (majorVer = 10) And (minorVer=0) Then
                                       Result := idOSWin10
                                     else
                                     Begin
//                                       ComplexVer := majorVer+(minorVer/10);
//
//                                       If ComplexVer>5.0  Then Result := idOSWin2000;
//                                       If ComplexVer>5.1  Then Result := idOSXP;
//                                       If ComplexVer>6.0  Then Result := idOSVista;
//                                       If ComplexVer>6.1  Then Result := idOSWin7;
//                                       If ComplexVer>6.2  Then Result := idOSWin8;
//                                       If ComplexVer>10.0 Then Result := idOSWin10;

                                       Result := idOSUnknown;
                                     end;
                                   end;
      VER_PLATFORM_WIN32_WINDOWS : begin { Windows 9x/ME }
                                     if (majorVer = 4) and (minorVer = 0) then
                                       Result := idOSWin95
                                     else if (majorVer = 4) and (minorVer = 10) then
                                     begin
                                       if osVerInfo.szCSDVersion[1] = 'A' then
                                         Result := idOSWin98SE
                                       else
                                         Result := idOSWin98;
                                     end else if (majorVer = 4) and (minorVer = 90) then
                                       Result := idOSWinME
                                     else
                                       Result := idOSUnknown;
                                   end;
      else
        Result := idOSUnknown;
    end;
  end
  else
    Result := idOSUnknown;
end;

function TDDUSystem.GetShellFolder(aShellFolder : TShellFolder): String;

Var
  R                     : TRegINIFile;

Begin
  R := TRegINIFile.Create('\Software\Microsoft\Windows\CurrentVersion\Explorer');
  Try
    Result := SafeAddBackSlash(R.ReadString('Shell Folders',ShellFolders[aShellFolder],''));
  Finally
    FreeAndNil(R);
  End;
End;

function TDDUSystem.GetShellFolderByName(aShellFolderName: String): String;

Var
  R                     : TRegINIFile;

Begin
  If SameText(aShellFolderName,'App_Path') Then
  Begin
    Result := ExtractFilePath(ParamStr(0));
  End
  Else
  Begin
    R := TRegINIFile.Create('\Software\Microsoft\Windows\CurrentVersion\Explorer');
    Try
      Result := R.ReadString('Shell Folders',aShellFolderName,'');
    Finally
      FreeAndNil(R);
    End;
  End;
  Result := SafeAddBackSlash(Result);
End;

function TDDUSystem.GetTempFile(Prefix: String): String;
begin
  SetLength(Result,MAX_PATH);
  If (GetTempFilename(PChar(fTempPath),PChar(Prefix),0,PChar(Result))<>0) Then
  Begin
    SetLength(Result,StrLen(PChar(Result)));
  End
  Else
  Begin
    Result := '';
  End;
end;

function TDDUSystem.GetTempFileAt(Dir, Prefix: String): String;
begin
  If Dir='' Then Dir := TempPath;
  CreatePath(Dir);
  SetLength(Result,MAX_PATH);
  If (GetTempFilename(PChar(Dir),PChar(Prefix),0,PChar(Result))<>0) Then
  Begin
    SetLength(Result,StrLen(PChar(Result)));
  End
  Else
  Begin
    Result := '';
  End;
end;

function TDDUSystem.GetTempFolder(Prefix: String): String;

begin
  Result := TempFolderAt['',Prefix];
end;

function TDDUSystem.GetTempFolderAt(Dir, Prefix: String): String;

begin
  Result := TempFileAt[Dir,Prefix];
  DeleteFile(Result);
  Result := Result+'\';
  CreatePath(Result,True);
end;

Function TDDUSystem.ShellDelete(Filename: String) : Boolean;

Var
  S : TStringList;

begin
  S := TStringList.Create;
  Try
    S.Add(FIleName);
    Result := ShellDelete(S);
  Finally
    FreeAndNil(S);
  End;
end;

Function TDDUSystem.ShellDelete(Files: TStrings) : Boolean;

Var
  Info                    : TSHFileOpStruct;
  S                       : String;
  S2                      : String;
  
begin
  S := StringReplace(Trim(Files.Text),#13#10,#0,[rfReplaceAll])+#0;
  S2 := #0#0;


  FillChar(Info,SizeOf(Info),0);

  Info.Wnd := 0;
  Info.wFunc := FO_DELETE;

  Info.pFrom := PChar(S);
  Info.pTo := PChar(S2);
  Info.fFlags := FOF_ALLOWUNDO or FOF_SILENT OR FOF_NOCONFIRMATION or FOF_NOERRORUI OR FOF_NOCONFIRMMKDIR;
  Info.fAnyOperationsAborted := False;

  Try
    Result := (SHFileOperation(Info)=0);
    Result := Result And (Not Info.fAnyOperationsAborted);
  Except
    Result := False;
  End;
end;

Function TDDUSystem.WinExec32At(FileName:String; aPath : String; Visibility : Integer) : DWord;

Var
  PaPath               : PChar;
  ProcessInfo           : TProcessInformation;
  StartupInfo           : TStartupInfo;

Begin
  Result := 0;
  If (aPath='') Then
  Begin
    PaPath := Nil;
  End
  Else
  Begin
    PaPath := PChar(aPath);
  End;
  FillChar(StartupInfo,Sizeof(StartupInfo),#0);
  StartupInfo.cb := Sizeof(StartupInfo);
  StartupInfo.dwFlags := STARTF_USESHOWWINDOW;
  StartupInfo.wShowWindow := Visibility;
  if not CreateProcess(nil,
    PChar(FileName),               { pointer to command line string }
    nil,                           { pointer to process security attributes }
    nil,                           { pointer to thread security attributes }
    false,                         { handle inheritance flag }
    CREATE_NEW_CONSOLE or          { creation flags }
    NORMAL_PRIORITY_CLASS,
    nil,                           { pointer to new environment block }
    PaPath,                       { pointer to current directory name }
    StartupInfo,                   { pointer to STARTUPINFO }
    ProcessInfo) then
  Begin
    Result := $ffffffff;
    MessageBox(0,PChar(Format('An error was encounted while trying to run:'#13'%s'#13'The error code is : %d',
                        [FileName,GetLastError])),'Error',mb_OK);
  End;
end;

function TDDUSystem.WinExecAndWait32(FileName:String; Visibility : Integer) : DWord;

Var
  WorkDir               : String;

begin
  GetDir(0,WorkDir);
  Result := WinExecAndWait32At(FileName,WorkDir,Visibility);
end;

Function TDDUSystem.WinExecAndWait32At(FileName : String; aPath : String; Visibility : Integer) : DWord;

Var
  PaPath                : PChar;
  ProcessInfo           : TProcessInformation;
  StartupInfo           : TStartupInfo;

Begin
  FillChar(StartupInfo,Sizeof(StartupInfo),#0);
  StartupInfo.cb := Sizeof(StartupInfo);
  StartupInfo.dwFlags := STARTF_USESHOWWINDOW;
  StartupInfo.wShowWindow := Visibility;
  If (aPath='') Then
  Begin
    PaPath := Nil;
  End
  Else
  Begin
    PaPath := PChar(aPath);
  End;

  if not CreateProcess(nil,
    PChar(FileName),               { pointer to command line string }
    nil,                           { pointer to process security attributes }
    nil,                           { pointer to thread security attributes }
    false,                         { handle inheritance flag }
    CREATE_NEW_CONSOLE or          { creation flags }
    NORMAL_PRIORITY_CLASS,
    nil,                           { pointer to new environment block }
    PaPath,                       { pointer to current directory name }
    StartupInfo,                   { pointer to STARTUPINFO }
    ProcessInfo) then
  Begin
    Result := High(Cardinal); { pointer to PROCESS_INF }
  End
  Else
  Begin
    WaitforSingleObject(ProcessInfo.hProcess,INFINITE);
    GetExitCodeProcess(ProcessInfo.hProcess,Result);
  End;
end;

Function TDDUSystem.WinExecEnv32(FileName:String; Env : String; Visibility : Integer) : DWord;

Var
  Loop                  : Integer;
  TEnv                  : String;
  ProcessInfo           : TProcessInformation;
  StartupInfo           : TStartupInfo;

Begin
  TEnv := AdjustLineBreaks(Env)+#0#0;
  For Loop := Length(TEnv) DownTo 1 Do
  Begin
    If TEnv[Loop]=#13 Then TEnv[Loop] := #0;
    If TEnv[Loop]=#10 Then Delete(TEnv,Loop,1);
  End;

  Result := 0;
  FillChar(StartupInfo,Sizeof(StartupInfo),#0);
  StartupInfo.cb := Sizeof(StartupInfo);
  StartupInfo.dwFlags := STARTF_USESHOWWINDOW;
  StartupInfo.wShowWindow := Visibility;
  if not CreateProcess(nil,PChar(FileName),
                       nil,nil,false,
                       CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS,
                       PChar(TEnv),
                       Nil,
                       StartupInfo,ProcessInfo) then
  Begin
    Result := $ffffffff;
    MessageBox(0,PChar(Format('An error was encounted while trying to run:'#13'%s'#13'The error code is : %d',
                        [FileName,GetLastError])),'Error',mb_OK);
  End;
end;


{ TDDUDesktop }

function TDDUDesktop.GetBottom: Integer;

Var
  R                       : TRect;

Begin
  SystemParametersInfo(SPI_GETWORKArea,0,@R,0);
  Result := R.Bottom;
end;

function TDDUDesktop.GetHeight: Integer;

Var
  R                       : TRect;

Begin
  SystemParametersInfo(SPI_GETWORKArea,0,@R,0);
  Result := R.Bottom-R.Top;
end;

function TDDUDesktop.GetLeft: Integer;

Var
  R                       : TRect;

Begin
  SystemParametersInfo(SPI_GETWORKArea,0,@R,0);
  Result := R.Left;
end;

function TDDUDesktop.GetPath: String;

Begin
  Result := DDUSystem.ShellFolder[sfDesktop];
  If (Result='') Then // For 95/98
  Begin
    Result := DDUSystem.WindowsPath+'Desktop\';
  End;
End;

function TDDUDesktop.GetRight: Integer;

Var
  R                       : TRect; 

Begin
  SystemParametersInfo(SPI_GETWORKArea,0,@R,0);
  Result := R.Right;
end;

function TDDUDesktop.GetTop: Integer;

Var
  R                       : TRect;

Begin
  SystemParametersInfo(SPI_GETWORKArea,0,@R,0);
  Result := R.Top;
end;

function TDDUDesktop.GetWidth: Integer;

Var
  R                       : TRect;

Begin
  SystemParametersInfo(SPI_GETWORKArea,0,@R,0);
  Result := R.Right-R.Left;
end;


{ TDDUEnvironment }

constructor TDDUEnvironment.Create;

Var
  Env                   : PChar;

begin
  fEnvironment := TStringList.Create;

  Env := GetEnvironmentStrings;
  While (Env[0]<>#0) DO
  Begin
    fEnvironment.Add(String(Env));
    Env := StrEnd(Env)+1;
  End;
  fEnvironment.Values[''] := '';
End;

destructor TDDUEnvironment.Destroy;
begin
  FreeAndNil(fEnvironment);
  inherited;
end;

function TDDUEnvironment.Expand(Source: String): String;

Var
  Index                   : Integer;
  At                      : Integer;
  At2                     : Integer;
  Len                     : Integer;
  aValue                  : String;
  Key                     : String;

Begin
  Index := 1;

  Result := '';
  Len := Length(Source);

  While Index<=Len Do
  Begin
    At := PosEx('%',Source,Index);

    If (At=Index) Then
    Begin
      At2 := PosEx('%',Source,At+1);
      If (At2=0) Then
      Begin
        Raise Exception.Create('Unmatched % at '+IntToStr(At));
      End;

      If (At2=At+1) Then
      Begin
        aValue := '%';
      End
      Else
      Begin
        Key := Copy(Source,At+1,At2-At-1);
        aValue := Value[Key];
      End;
      Result := Result+aValue;

      Index := At2+1;
    End
    Else
    Begin
      If (At=0) Then
      Begin
        At := Len+1;
      End;
      Result := Result+Copy(Source,Index,At-Index);
      Index := At;
    End;
  End;
End;

function TDDUEnvironment.GetExists(Name: String): Boolean;
begin
  Result := (fEnvironment.IndexOfName(Name)<>-1);
end;

function TDDUEnvironment.GetText: String;
begin
  Result := fEnvironment.Text;
end;

function TDDUEnvironment.GetValue(Name: String): String;
begin
  Result := fEnvironment.Values[Name];
end;

function TDDUEnvironment.GetValueDef(Name, aDefault: String): String;

Var
  At                      : Integer;

Begin
  At := fEnvironment.IndexOfName(Name);
  If At=-1 Then
  Begin
    Result := aDefault;
  End
  Else
  Begin
    Result := fEnvironment.ValueFromIndex[At];
  End
End;

procedure TDDUEnvironment.LoadFromStrings(Source: TStrings);
begin
  fEnvironment.Assign(Source);
end;

procedure TDDUEnvironment.LoadFromSystem;
Var
  Env                   : PChar;

begin
  fEnvironment.Clear;
  Env := GetEnvironmentStrings;
  While (Env[0]<>#0) DO
  Begin
    fEnvironment.Add(String(Env));
    Env := StrEnd(Env)+1;
  End;
  fEnvironment.Values[''] := '';
End;


procedure TDDUEnvironment.SetValue(Name: String; const Value: String);
begin
  fEnvironment.Values[Name] := Value;
end;

{ TDDUPath }

function TDDUPath.AddBackslash(const aPath : String): String;

begin
  If (Copy(aPath,Length(aPath),1)<>'\') Then
  Begin
    Result := aPath+'\';
  End
  Else
  Begin
    Result := aPath;
  End;
end;

procedure TDDUPath.CopyTree(aSourcePath, aDestPath: String);

Var                     
  DosError                : Integer;    
  SearchRec               : TSearchRec; 

Begin
  aSourcePath := SafeAddBackslash(aSourcePath);
  aDestPath   := SafeAddBackslash(aDestPath);

  If (aSourcePath='') Then
  Begin
    Raise CopyTreeException.Create('No source.');
  End;
  If (aDestPath='') Then
  Begin
    Raise CopyTreeException.Create('No destination.');
  End;

  If IsRoot(aSourcePath) Then
  Begin
    Raise CopyTreeException.CreateFmt('Source is root : %s',[aSourcePath]);
  End;

  If Not CreatePath(aDestPath) Then
  Begin
    Raise Exception.CreateFmt('Can not create destination : %s',[aDestPath]);
  End;

  DosError := FindFirst(aSourcePath+'*.*',faAnyFile,SearchRec);
  Try
    While (DosError=0) Do
    Begin
      If (SearchRec.Attr and faVolumeID)<>0 Then
      Begin
        // Do nothing.
      End Else If ((SearchRec.Attr And faDirectory)<>0) Then
      Begin
       If ((SearchRec.Name<>'.') And (SearchRec.Name<>'..')) Then
       Begin
         CopyTree(aSourcePath+SearchRec.Name+'\',aDestPath+SearchRec.Name+'\');
       End;
      End Else
      Begin
        If ((SearchRec.Attr And (faReadOnly or faHidden))<>0) Then
        Begin
          FileSetAttr(aSourcePath+SearchRec.Name,0)
        End;

        If Not CopyFile(PChar(aSourcePath+SearchRec.Name), PChar(aDestPath+SearchRec.Name), False) Then
        Begin
          Raise CopyTreeException.CreateFmt('Could not copy (%d)%s : %s to %s',[GetLastError,SysErrorMessage(GetLastError),aSourcePath+SearchRec.Name,aDestPath+SearchRec.Name ]);
        End;
      End;
      DosError := FindNext(SearchRec);
    End;
  Finally
    System.SysUtils.FindClose(SearchRec);
  End;
End;

function TDDUPath.CreatePath(const APath: String): Boolean;

Var
  FullName              : String;
  Work                  : String;
  Frag                  : String;

Begin
  FullName := AddBackSlash(ExpandUNCFileName(APath));
  Work := AddBackSlash(ExtractFileDrive(FullName));
  Repeat
    Frag := Copy(FullName,Length(Work)+1,Length(FullName));  // Get everything yet to be created.
    If Frag<>'' Then
    Begin
      Delete(Frag,Pos('\',Frag)+1,Length(Frag));         // We only need the next directory.
      Work := Work+Frag;
      CreateDir(Work);
    End;
  Until (Frag='');
  Result := DirExists(APath);
end;

function TDDUPath.DirExists(aPath : string): Boolean;

Var
  Code                  : Integer;

begin
  Code := GetFileAttributes(PChar(aPath));
  Result := (Code <> -1) and (FILE_ATTRIBUTE_DIRECTORY and Code <> 0);
end;

function TDDUPath.FullPathToRelativePath(BasePath, FullPath: String): String;
begin
  FullPath := ExtractRelativePath(BasePath,FullPath);
  If (Pos(':',FullPath)=0) And (Pos('\\',FullPath)<>1) Then
  Begin
    Result := '~\'+FullPath;
  End
  Else
  Begin
    Result := FullPath;
  End;
end;

function TDDUPath.IsRoot(aPath: String): Boolean;

Var
  Work                    : String;

begin
  aPath := AddBackSlash(ExpandUNCFileName(aPath));
  Work := ExtractFileDrive(aPath);
  Result := (Copy(aPath,Length(Work)+1,Length(aPath))='\');
end;

procedure TDDUPath.KillTree(aPath: String);

Var                     
  SearchRec               : TSearchRec; 
  DosError                : Integer;    

Begin
  aPath := SafeAddBackslash(aPath);
  If (AnsiCompareFileName(SafeAddBackslash(GetCurrentDir),aPath)=0) Then
  Begin
    ChDir(ParentDir(GetCurrentDir));
  End;

  If (aPath<>'') And (Not IsRoot(aPath)) And DirExists(aPath) Then
  Begin
    DosError := FindFirst(aPath+'*.*',faAnyFile,SearchRec);
    While (DosError=0) Do
    Begin
      If ((SearchRec.Attr And faVolumeID)=faVolumeID) Then
      Begin
      End Else If ((SearchRec.Attr And faDirectory)=faDirectory) Then
      Begin
        If (SearchRec.Name<>'.') And (SearchRec.Name<>'..') Then
        Begin
          KillTree(aPath+SearchRec.Name);
        End;
      End Else
      Begin
        If ((SearchRec.Attr And (faReadOnly or faHidden))<>0) Then
        Begin
          FileSetAttr(aPath+SearchRec.Name,0)
        End;
        System.SysUtils.DeleteFile(aPath+SearchRec.Name);
      End;
      DosError := FindNext(SearchRec);
    End;
    System.SysUtils.FindClose(SearchRec);

    Try
      RmDir(aPath);
    Except
    End;
  End;
End;

function TDDUPath.LoseBackslash(const aPath: String): String;
begin
  Result := aPath;
  If Not IsRoot(aPath) Then
  Begin
    If Copy(aPath,Length(aPath),1)='\' Then
    Begin
      SetLength(Result,Length(Result)-1);
    End;
  End;
end;

Procedure TDDUPath.MoveTree(aSourcePath : String; aDestPath : String);

Begin
  aSourcePath := SafeAddBackslash(aSourcePath);
  aDestPath   := SafeAddBackslash(aDestPath);

  If (aSourcePath='') Then
  Begin
    Raise CopyTreeException.Create('No source.');
  End;
  If (aDestPath='') Then
  Begin
    Raise CopyTreeException.Create('No destination.');
  End;
  If IsRoot(aSourcePath) Then
  Begin
    Raise CopyTreeException.CreateFmt('Source is root : %s',[aSourcePath]);
  End;

  If (AnsiCompareFileName(SafeAddBackslash(GetCurrentDir),aSourcePath)=0) Then
  Begin
    ChDir(ParentDir(GetCurrentDir));
  End;

  Try
    CopyTree(aSourcePath,aDestPath);
  Except
    On E:Exception Do
    Begin
      KillTree(aDestPath);
      Raise MoveTreeException.Create(E.Message);
    End;
  End;
  KillTree(aSourcePath);
End;

function TDDUPath.ParentDir(const aPath: String): String;
begin
  Result := AddBackslash(aPath);
  If Not IsRoot(Result) Then
  Begin
    Result := ExtractFilePath(LoseBackslash(Result))
  End;
end;

function TDDUPath.RelativePathToFullPath(BasePath, RelativePath: String): String;
begin
  If RelativePath='~' Then
  Begin
    Result := BasePath;
  End Else If Copy(RelativePath,1,2)='~\' Then
  Begin
    Result := BasePath+Copy(RelativePath,3,Length(RelativePath)-2);
  End Else If Copy(RelativePath,1,1)='~' Then
  Begin
    Result := BasePath+Copy(RelativePath,2,Length(RelativePath)-1);
  End Else
  Begin
    Result := RelativePath;
  End;

  If (Copy(BasePath,1,1)<>'~') Then
  Begin
    Result := ExpandUNCFileName(Result);
  End;
end;

function TDDUPath.SafeAddBackslash(const aPath: String): String;
begin
  If (aPath='') Then
  Begin
    Result := '';
  End
  Else
  Begin
    Result := Addbackslash(aPath);
  End;
end;

procedure TDDUPath.WalkTree(aPath, aMask: String; Recurse: Boolean; OnFile, OnFolder: TOnWalkTreeEvent);

Var
  SearchRec               : TSearchRec;
  DosError                : Integer;

Begin
  aPath := SafeAddBackSlash(aPath);

  DosError := FindFirst(aPath+'*.*',faAnyFile,SearchRec);
  Try
    While (DosError=0) Do
    Begin
      If (SearchRec.Attr and faVolumeID)<>0 Then
      Begin
        // Do nothing.
      End Else If (SearchRec.Attr and faDirectory)<>0 Then
      Begin
        If (SearchRec.Name<>'.') And (SearchRec.Name<>'..') Then
        Begin
          If Assigned(onFolder) Then
          Begin
            OnFolder(aPath+SearchRec.Name+'\',SearchRec);
          End;
          If Recurse Then
          Begin
            WalkTree(aPath+SearchRec.Name+'\',aMask,True,OnFile,OnFolder);
          End;
        End;
      End
      Else
      Begin
        If Assigned(OnFile) And DDUMatchesMask(SearchRec.Name,aMask) Then
        Begin
          OnFile(aPath+SearchRec.Name,SearchRec);
        End;
      End;
      DosError := FindNext(SearchRec);
    End;
  Finally
    FindClose(SearchRec);
  End;
End;

{ TDDURTTI }

function TDDURTTI.CharSetToString(C: TCharSet): String;

Var
  Loop                  : Char;
  Low                   : Char;
  High                  : Char;
  Valid                 : Boolean;

Begin
  Result := '';

  Low := #0;
  High := #0;
  Valid := False;

  For Loop := #0 To #255 Do
  Begin
    If CharInSet(Loop,C) Then
    Begin
      If Not Valid Then
      Begin
        Low := Loop;
      End;
      High := Loop;
      Valid := True;
    End
    Else
    Begin
      If Valid Then Result := Result+Emit(Low,High);
      Valid := False;
    End;
  End;

  If Valid Then Result := Result+Emit(Low,High);

  If Length(Result)>0 THen SetLength(Result,Length(Result)-1);
End;

procedure TDDURTTI.CreateSet(Source: String; var ResultSet: TCharSet);

Var
  At                    : Integer;

function GetToken(IsLitteral : Boolean) : Char;

Begin
  If (At<=Length(Source)) Then
  Begin
    Result := Source[At];
    Inc(at);
    If (Not IsLitteral) Then
    Begin
      If (Result='\') Then  // Litteral character.
      Begin
        Result := GetToken(True);
      End;
// Provides support for embeded control characters with the standard Pascal ^a format.
// Removed because ^ has a special meaning for scanf sets.  Kept for future reference.
(*      Else If (Result='^') Then
      Begin
        Result := UpCase(GetToken(True));
        If Not (Result in [#64..#95]) Then
        Begin
          Raise Exception.Create('BAD SET : '+Source);
        End;
        Result := Char(Byte(Result)-64);
      End; *)
    End;
  End
  Else
  Begin
    Raise Exception.Create('BAD SET : '+Source);
  End;
End;

Var
  Token                 : Char;
  EndToken              : Char;
  LoopToken             : Char;
  Negate                : Boolean;

Begin
  ResultSet := [];
  At := 1;
  Negate := (Copy(Source,1,1)='^');
  If Negate Then
  Begin
    Delete(Source,1,1);
    ResultSet := [#0..#255];
  End;
  While (At<=Length(Source)) Do
  Begin
    Token := GetToken(False);
    EndToken := Token;

    If (Copy(Source,At,1)='-') And (At<>Length(Source)) Then  // This is a range.
    Begin
      Inc(At,1);  // Go past dash.

      If At>Length(Source) Then
      Begin
        Raise Exception.Create('BAD SET : '+Source);
      End;
      EndToken := GetToken(False);
    End;

    If (EndToken<Token) Then  // Z-A is probably the result of a missing letter.
    Begin
      Raise Exception.Create('BAD SET : '+Source);
    End;

    For LoopToken := Token To EndToken Do
    Begin
      Case Negate of
        False : Include(ResultSet,AnsiChar(LoopToken));
        True  : Exclude(ResultSet,AnsiChar(LoopToken));
      End;
    End;
  End;
end;

function TDDURTTI.Emit(Low, High: Char): String;
begin
  If (Low=High) Then
  Begin
    Result := EmitChar(Low)+',';
  End
  Else
  Begin
    Result := EmitChar(Low)+'..'+EmitChar(High)+',';
  End;
end;

function TDDURTTI.EmitChar(What: Char): String;
begin
  Case What Of
    #0..#32,
    #127..#255 : Result := '#'+IntToStr(Ord(What));
  Else
    Result := What;
  End;
end;

function TDDURTTI.EnumToString(Const Value; Info: PTypeInfo): String;

Var
  Data                    : PTypeData;

Begin
  Data := GetTypeData(Info);
  If (Cardinal(Value)<Cardinal(Data^.MinValue)) Or (Cardinal(Value)>Cardinal(Data^.MaxValue)) Then
  Begin
    Result := Format('%s(%d)',[Info^.Name,Cardinal(Value)]);
  End
  Else
  Begin
    Result := GetEnumName(Info,Cardinal(Value));
  End;
end;

function TDDURTTI.ReverseEmit(What: String): Char;
begin
  If Length(What)=1 THen
  Begin
    Result := What[1];
  End
  Else
  Begin
    Result := Char( StrToInt(Copy(What,2,Length(What)-1)) );
  End;
end;

function TDDURTTI.SetToString(Const SetValue; Info: PTypeInfo): String;

Var
  CompData                : PTypeData;
  CompInfo                : PTypeInfo;
  Data                    : PTypeData;
  Loop                    : Integer;
  Mask                    : Integer;
  Value                   : PByte;

Begin
  Data := GetTypeData(Info);
  CompInfo :=  Data^.CompType^;

  CompData := GetTypeData(CompInfo);

  Value  := @SetValue;
  Mask   := 1;
  Result := '[';

  For Loop := CompData^.MinValue To CompData^.MaxValue Do
  Begin
    If (Value^ And Byte(Mask))<>0  Then
    Begin
      Result := Result+EnumToString(Loop,CompInfo)+',';
    End;
    Mask := Mask Shl 1;
    If Mask=$100 Then
    Begin
      Inc(Value);
      Mask := 1;
    End;
  End;
  If (Result<>'[') Then
  Begin
    SetLength(Result,Length(Result)-1);
  End;
  Result := Result+']';
end;

procedure TDDURTTI.StringToCharSet(S: String; var C: TCharSet);

Var
  At                    : Integer;
  Frag                  : String;
  Low,High              : Char;
  Loop                  : Char;

Begin
  C := [];
  While Length(S)<>0 Do
  Begin
    At := Pos(',',Copy(S,2,Length(S)-1));
    If At=0 Then
    Begin
      Frag := S;
      S := ''
    End
    Else
    begin
      Inc(at);
      Frag := Copy(S,1,At-1);
      Delete(S,1,At);
    End;

    At := Pos('..',Frag);
    If At=0 Then // Single Character
    Begin
      Include(C,AnsiChar(ReverseEmit(Frag)));
    End
    Else
    Begin
      Low := ReverseEmit(Copy(Frag,1,At-1));
      Delete(Frag,1,At+1);
      High := ReverseEmit(Frag);
      For Loop := Low To High Do
      Begin
        Include(C,AnsiChar(Loop));
      End;
    End;
  End;
End;

function TDDURTTI.StringToEnum(Value: String; Info: PTypeInfo): Integer;
begin
  Result := GetEnumValue(Info,Value);
end;

procedure TDDURTTI.StringToSet(S: String; Info: PTypeInfo; var SetValue);

Var
  CompData                : PTypeData;
  CompInfo                : PTypeInfo;
  Data                    : PTypeData;
  Items                   : TStringList;
  Loop                    : Integer;
  Mask                    : Integer;
  Result                  : PByte;

Begin
  Data := GetTypeData(Info);
  CompInfo :=  Data^.CompType^;

  CompData := GetTypeData(CompInfo);

  Result := @SetValue;
  Mask := 1;

  Items := TStringList.Create;
  Try
    If Copy(S,1,1)='[' Then Delete(S,1,1);
    If Copy(S,Length(S),1)=']' Then Delete(S,Length(S),1);

    Items.Text := StringReplace(S,',',#13#10,[rfReplaceAll]);
    For Loop := Items.Count-1 DownTo 0 Do
    Begin
      Items[Loop] := Trim(Items[Loop]);
      If Items[Loop]='' Then
      Begin
        Items.Delete(Loop);
      End;
    End;

    For Loop := CompData^.MinValue To CompData^.MaxValue Do
    Begin
      If Items.IndexOf( EnumToString(Loop,CompInfo) )<>-1 Then
      Begin
        Result^ := Result^ Or Byte(Mask);
      End
      Else
      Begin
        Result^ := Result^ And (Not Byte(Mask));
      End;
      Mask := Mask Shl 1;
      If Mask=$100 Then
      Begin
        Mask := 1;
        Inc(Result);
      End;
    End;
  Finally
    FreeAndNil(items);
  End;
end;

Procedure AspectRatio(DisplayW, DisplayH : Integer; SourceW,SourceH : Extended; Var FinalW,FinalH : Integer);

Begin
  FinalW := DisplayW;
  FinalH := Round(SourceH*DisplayW/SourceW);

  If FinalH>DisplayH Then
  Begin
    FinalH := DisplayH;
    FinalW := Round(SourceW*DisplayH/SourceH);
  End;
End;

Procedure AspectRatio(DisplayW, DisplayH : Extended; SourceW,SourceH : Extended; Var FinalW,FinalH : Extended);

Begin
  FinalW := DisplayW;
  FinalH := SourceH*DisplayW/SourceW;

  If FinalH>DisplayH Then
  Begin
    FinalH := DisplayH;
    FinalW := SourceW*DisplayH/SourceH;
  End;
End;

Procedure ScaleERect(SourceRect : TERect; DestRect : TERect; Input : TERect; Var ScaledResult : TERect);

Var
  RatioX                  : Extended;
  RatioY                  : Extended;

Begin
  RatioX := (DestRect.Right-DestRect.Left) / (SourceRect.Right- SourceRect.Left);
  RatioY := (DestRect.Top-DestRect.Bottom) / (SourceRect.Top- SourceRect.Bottom);

  ScaledResult.Left   := (( Input.Left-SourceRect.Left)    *RatioX) +DestRect.Left;   
  ScaledResult.Bottom := ((Input.Bottom-SourceRect.Bottom) *RatioY) +DestRect.Bottom; 
  ScaledResult.Right  := ((Input.Right-SourceRect.Left)    *RatioX) +DestRect.Left;   
  ScaledResult.Top    := ((Input.Top-SourceRect.Bottom)    *RatioY) +DestRect.Bottom;
End;

{ TDDUEScale }

constructor TDDUEScale.Create(aWorld, aScaledWorld: TERect);
begin
  Inherited Create;
  fWorld := aWorld;
  fScaledWorld := aScaledWorld;
  fRatioX := (fScaledWorld.Right-fScaledWorld.Left) / (fWorld.Right-fWorld.Left);
  fRatioY := (fScaledWorld.Top-fScaledWorld.Bottom) / (fWorld.Top-fWorld.Bottom);
end;

function TDDUEScale.Scale(R: TERect): TERect;
begin
  Result.Left   := (( R.Left-fWorld.Left)    *fRatioX) +fScaledWorld.Left;
  Result.Bottom := ((R.Bottom-fWorld.Bottom) *fRatioY) +fScaledWorld.Bottom;
  Result.Right  := ((R.Right-fWorld.Left)    *fRatioX) +fScaledWorld.Left;
  Result.Top    := ((R.Top-fWorld.Bottom)    *fRatioY) +fScaledWorld.Bottom;
end;

{ TDDUIntScale }

constructor TDDUIntScale.Create(aWorld, aScaledWorld: TRect);
begin
  Inherited Create;
  fWorld       := aWorld;       
  fScaledWorld := aScaledWorld; 
  fRatioX := (fScaledWorld.Right-fScaledWorld.Left) / (fWorld.Right-fWorld.Left);
  fRatioY := (fScaledWorld.Top-fScaledWorld.Bottom) / (fWorld.Top-fWorld.Bottom);
end;

function TDDUIntScale.Scale(R: TRect): TRect;
begin
  Result.Left   := Round((( R.Left-fWorld.Left)    *fRatioX) +fScaledWorld.Left);
  Result.Bottom := Round(((R.Bottom-fWorld.Bottom) *fRatioY) +fScaledWorld.Bottom);
  Result.Right  := Round(((R.Right-fWorld.Left)    *fRatioX) +fScaledWorld.Left);
  Result.Top    := Round(((R.Top-fWorld.Bottom)    *fRatioY) +fScaledWorld.Bottom);
end;

{ TDDUFolderEnumerator }

constructor TDDUFolderEnumerator.Create(aSearch : String; aAttr :Integer);
begin
  Inherited Create;
  fAttr := aAttr;
  fSearch := aSearch;
  fStarted := False;
end;

destructor TDDUFolderEnumerator.Destroy;
begin
  If fStarted Then
  Begin
    FindClose(fSearchRec);
  End;
end;

function TDDUFolderEnumerator.GetCurrent: TSearchRec;
begin
  Result := fSearchRec;
end;

function TDDUFolderEnumerator.MoveNext: Boolean;

Var
  DosError                : Integer;
begin
  If Not fStarted Then
  Begin
    DosError := FindFirst(fSearch,fAttr,fSearchRec);
    fStarted := True;
  End
  Else
  Begin
    DosError := FindNext(fSearchRec);
  End;

  Result := (DosError=0);
end;

{ TDDUFolder }

constructor TDDUFolder.Create(aFolder, aMask: String; aAttr : Integer);
begin
  Inherited Create;
  fAttr := aAttr;
  fFolder := aFolder;
  fMask   := aMask;
end;

function TDDUFolder.GetEnumerator: TDDUFolderEnumerator;
begin
  Result := TDDUFolderEnumerator.Create( SafeAddBackSlash(fFolder)+fMask,fAttr );
end;

Function ERectToRect(ERect : TERect) : TRect;

Begin
  Result.Left := Round(ERect.Left);
  Result.Bottom := Round(ERect.Bottom);
  Result.Right:= Round(ERect.Right);
  Result.Top:= Round(ERect.Top);
End;

Function RectToERect(Rect : TRect) : TERect;

Begin
  Result.Left := Rect.Left;
  Result.Bottom := Rect.Bottom;
  Result.Right:= Rect.Right;
  Result.Top:= Rect.Top;
End;

Function ERect(Left,Bottom,Right,Top : Extended) : TERect;

Begin
  Result.Left := Left;
  Result.Bottom := Bottom;
  Result.Right:= Right;
  Result.Top:= Top;
End;

Function KB_Value(Value : UInt64) : String;

Const
  Scale                   : Array[0..4] Of String =('B','KB', 'MB','GB','TB');

Var
  V                       : Extended;
  sIndex                  : Integer;

Begin
  sIndex := 0;
  V := Value;

  While sIndex<High(Scale) Do
  Begin
    If V<1024 Then
    Begin
      Break;
//      Result := Format('%0.1f %s',[V,Scale[sIndex]]);
//      Exit;
    End;

    V := V/1024;
    Inc(sIndex);
  End;
  Result := Format('%0.1f %s',[V,Scale[sIndex]]);
End;


Function PrintableTime(Milliseconds : Cardinal) : String;


Var
  Days                    : Cardinal;
  Hours                   : Cardinal;
  Minutes                 : Cardinal;
  Seconds                 : Cardinal;

Begin

  Days         := Trunc(Milliseconds/ MillisecondsPerDay);
  Milliseconds := Milliseconds-Days*MillisecondsPerDay;
  Hours        := Trunc(Milliseconds/MillisecondsPerHour);
  Milliseconds := Milliseconds-Hours*MillisecondsPerHour;
  Minutes      := Trunc(Milliseconds/MillisecondsPerMinute);
  Milliseconds := Milliseconds-Minutes*MillisecondsPerMinute;
  Seconds      := Trunc(Milliseconds/MillisecondsPerSecond);
  Milliseconds := Milliseconds-Seconds*MillisecondsPerSecond;

  Result := Format('%d ms',[Milliseconds]);
  If (Seconds<>0) Or (Minutes<>0) Or (Hours<>0) Or (Days<>0) Then
  Begin
    Result := Format('%.2d:%.2d, %s',[Minutes,Seconds,Result]);
  End;
  If (Hours<>0) Or (Days<>0) Then
  Begin
    REsult := Format('%.2d:%s',[Hours,Result]);
  End;
  If (Days<>0) Then
  Begin
    REsult := Format('%d days, %s',[Hours,Result]);
  End;
End;

Function PrintableTimeEx(Milliseconds : Cardinal) : String;


Var
  Days                    : Cardinal;
  Hours                   : Cardinal;
  Minutes                 : Cardinal;
  Seconds                 : Cardinal;

Begin

  Days         := Trunc(Milliseconds/ MillisecondsPerDay);
  Milliseconds := Milliseconds-Days*MillisecondsPerDay;
  Hours        := Trunc(Milliseconds/MillisecondsPerHour);
  Milliseconds := Milliseconds-Hours*MillisecondsPerHour;
  Minutes      := Trunc(Milliseconds/MillisecondsPerMinute);
  Milliseconds := Milliseconds-Minutes*MillisecondsPerMinute;
  Seconds      := Trunc(Milliseconds/MillisecondsPerSecond);
  Milliseconds := Milliseconds-Seconds*MillisecondsPerSecond;

  If Milliseconds<>0 Then
  Begin
    Result := Format('%d ms',[Milliseconds]);
  End
  Else
  Begin
    Result := '';
  End;
  If (Seconds<>0) Then
  Begin
    Result := Trim(Format('%d seconds %s',[Seconds,Result]));
  End;
  If (Minutes<>0) Then
  Begin
    Result := Trim(Format('%d Minutes %s',[Minutes,Result]));
  End;
  If (Hours<>0) Then
  Begin
    Result := Trim(Format('%d Hours %s',[hours,Result]));
  End;
  If (Days<>0) Then
  Begin
    Result := Trim(Format('%d Days %s',[Days,Result]));
  End;
  If Result='' Then
  Begin
    Result := '0 ms';
  End;
End;

Function PrintableTime(aTime : TDateTime; ShowMS: Boolean=True) : String;

Var
  Days                    : Cardinal;
  Hours                   : Cardinal;
  Minutes                 : Cardinal;
  Seconds                 : Cardinal;
  Milliseconds            : Cardinal;

  H,M,S,MS                : Word;

  Part1,Part2,Part3       : String;

Begin
  DecodeTime(aTime,H,M,S,MS);

  Days         := Trunc(aTime);
  Hours        := H;
  Minutes      := M;
  Seconds      := S;
  Milliseconds := MS;

  If (Days<>0) Then
  Begin
    Part1 := Format('%d d',[Days]);
  End
  Else
  Begin
    Part1 := '';
  End;

  If (Hours<>0) Then
  Begin
    Part2 := Format('%.2d:%.2d:%.2d',[Hours,Minutes,Seconds]);
  End Else If (Minutes<>0) Or (Seconds<>0) Then
  Begin
    Part2 := Format('%.2d:%.2d',[Minutes,Seconds]);
  End Else
  Begin
    Part2 := '';
  End;

  If (Milliseconds<>0) And ShowMS Then
  Begin
    Part3 := Format('%d ms',[Milliseconds]);
  End
  Else
  Begin
    Part3 := '';
  End;
  Result := Part1;
  If (Part1<>'') And ((Part2<>'') Or (Part3<>'')) Then
  Begin
    Result := Result+', ';
  End;
  Result := Result+Part2;
  If (Part2<>'') And (Part3<>'') Then
  Begin
    Result := Result+', ';
  End;
  Result := Result+Part3;
End;

Function PrintableTimeEx(aTime : TDateTime; ShowMS: Boolean=True) : String;

Var
  Days                    : Cardinal;
  Hours                   : Cardinal;
  Minutes                 : Cardinal;
  Seconds                 : Cardinal;
  Milliseconds            : Cardinal;

  H,M,S,MS                : Word;

//  Part1,Part2,Part3       : String;

Begin
  DecodeTime(aTime,H,M,S,MS);

  Days         := Trunc(aTime);
  Hours        := H;
  Minutes      := M;
  Seconds      := S;
  Milliseconds := MS;

  If (Milliseconds<>0) And ShowMS Then
  Begin
    Result := Format('%d ms',[Milliseconds]);
  End
  Else
  Begin
    Result := '';
  End;
  If (Seconds<>0) Then
  Begin
    Result := Trim(Format('%d seconds %s',[Seconds,Result]));
  End;
  If (Minutes<>0) Then
  Begin
    Result := Trim(Format('%d Minutes %s',[Minutes,Result]));
  End;
  If (Hours<>0) Then
  Begin
    Result := Trim(Format('%d Hours %s',[hours,Result]));
  End;
  If (Days<>0) Then
  Begin
    Result := Trim(Format('%d Days %s',[Days,Result]));
  End;
  If Result='' Then
  Begin
    Result := '0 ms';
  End;
End;

Function NewGUIDString : String;

Var
  GUID                    : TGUID;

begin
  CreateGUID(GUID);
  Result := GUIDToString(GUID);
End;

procedure SetCurrentThreadName(const Name: string);

Type
  TThreadNameInfo = Record
                      RecType: LongWord;
                      Name: PChar;
                      ThreadID: LongWord;
                      Flags: LongWord;
                    end;

var
  info                    : TThreadNameInfo;
   
begin
  // This code is extremely strange, but it's the documented way of doing it!
  info.RecType  := $1000;       
  info.Name     := PChar(Name); 
  info.ThreadID := $FFFFFFFF;   
  info.Flags    := 0;           
  try
    RaiseException($406D1388, 0,       SizeOf(info) div SizeOf(LongWord), PUINT_PTR(@info));
  except
  end;
end;

Procedure FastHex(B : Byte; Var C1,C2 : AnsiChar); Inline;

Var
  B1,B2 : Integer;

Begin
  B1 := (B And $f0) Shr 4;
  B2 := B and $f;

  If B1>=10 Then C1 := AnsiChar(Ord(B1+Ord('A')-10)) Else C1 := AnsiChar(Ord(B1+Ord('0')));
  If B2>=10 Then C2 := AnsiChar(Ord(B2+Ord('A')-10)) Else C2 := AnsiChar(Ord(B2+Ord('0')));
End;

Function FastUnHex(C : AnsiChar) : Byte; Inline;

Begin
  Case C Of
    '0'..'9' : Result := Ord(C)-Ord('0');
    'A'..'F' : Result := Ord(C)-Ord('A')+10;
    'a'..'f' : Result := Ord(C)-Ord('a')+10;
  Else
    Raise Exception.Create('Corrupted hex');
  End;
End;

//function Pos(const substr, str: AnsiString): Integer; overload;
function RightPos(const substr, str: String): Integer;

Var
  Loop                    : Integer;

Begin
  Result := 0;

  For Loop := Length(Str)-Length(SubStr)+1 Downto 1 Do
  Begin
    If Copy(Str,Loop,Length(SubStr))=SubStr Then
    Begin
      Result := Loop;
      Break;
    End;
  End;
End;

Function FastXMLFormat(XMLString: String): String;

Var
  Loop                    : Integer;
  SubLoop                 : Integer;
  DataLen                 : Integer;
  Tags                    : TStringList;
  Data                    : String;
  Indent                  : Integer;
  S2                      : TStringList;

Begin
  Tags := TStringList.Create;
  Try
    Loop := 1;
    While Loop<=Length(XMLString) Do
    Begin
      If (XMLString[Loop]='<') Then
      Begin
        DataLen := 1;
        While (XMLString[Loop+DataLen]<>'>') And (Loop+DataLen<LEngth(XMLString)) Do Inc(DataLen);
        If (XMLString[Loop+DataLen]='>')  Then Inc(DataLen);
      End
      Else
      Begin
        DataLen := 1;
        While (XMLString[Loop+DataLen]<>'<') And (Loop+DataLen<LEngth(XMLString)) Do Inc(DataLen);
      End;
      Data := Copy(XMLString,Loop,DataLen);
      If Trim(Data)<>'' Then
      Begin
        Tags.Add(Data);
      End;
      Loop := Loop+DataLen;
    End;
    Indent := 0;

    For Loop := 0 To Tags.Count-1 Do
    Begin
      Data := Trim(Tags[Loop]);

      If (Copy(Data,1,1)='<') Then
      Begin
        If (Copy(Data,1,2)='</') Then
        Begin
          Dec(Indent);
          If Indent<0 Then
          Begin
            Indent := 0;
          End;
          Tags[Loop] := StringOfChar(#32,Indent*2)+Trim(Data);
        End Else If (Copy(Data,Length(Data)-1,2)='/>') Then
        Begin
          // Self contained tag.
          Tags[Loop] := StringOfChar(#32,Indent*2)+Trim(Data);
        End
        Else
        Begin
          Tags[Loop] := StringOfChar(#32,Indent*2)+Trim(Data);
          Inc(Indent);
        End;
      End
      Else
      Begin
        S2 := TStringList.Create;
        Try
          S2.Text := Data;

          For SubLoop := 0 To S2.Count-1 Do
          Begin
            S2[SubLoop] := StringOfChar(#32,Indent*2)+Trim(S2[SubLoop]);
          End;

          Data := TrimRight(S2.Text);
        Finally
          FreeAndNil(S2);
        End;
        Tags[Loop] := Data;
      End;
    End;
    Result := Tags.Text;
  Finally
    FreeAndNil(Tags);
  End;
End;

function DDUMatchesMask(const TestSample, Mask: string): Boolean;

Begin
  Result := MatchesMask(Testsample,StringReplace(Mask,'[','[[]',[rfReplaceAll]));
End;

Procedure GetStringPropertyList(aObj : TObject; Names : TStrings);

Var
  Info                    : PTypeInfo;
  Data                    : PTypeData;
  PropList                : PPropList;
  Count                   : Integer;
  I                       : Integer;
  PropName                : String;
  PropInfo                : PPropInfo;
//  aPropInfo               : PPropInfo;
  O                       : TObject;

Begin
  Info := aObj.ClassInfo;
  Data := GetTypeData(Info);
  GetMem(PropList,Data.PropCount*SizeOf(PPropInfo));
  Try
    Count := GetPropList(info,tkAny ,propList);
    For I := 0 To Count-1 Do
    Begin
      PropInfo := PropList^[I];
      PropName := String(PropInfo.Name);

      Case PropInfo^.PropType^.Kind Of
        tkLString,
        tkWString,
        tkString : Names.Add(PropName);
        tkClass  : Begin
                     If Assigned(PropInfo) Then
                     Begin
                       O := TObject( GetOrdProp(aObj,PropInfo) );
                       If Assigned(O) And (O Is TStrings) Then
                         Names.Add(PropName);
                     End;
                   End;
      Else
//        Names.Add(EnumToString(PropInfo^.PropType^.Kind,TypeInfo(TTypeKind))+' '+PropName);
      End;
    End;
  Finally
    FreeMem(PropList);
  End;
End;

Procedure MySetStrProp(Instance: TObject; const PropName: string; const Value: string);

Var
  PropInfo : PPropInfo;
  O        : TObject;

Begin
  PropInfo := GetPropInfo(Instance.ClassInfo,PropName);

  If Assigned(PropInfo) Then
  Begin
    Case (PropInfo^.PropType^.Kind) Of
      tkLString,
      tkWString,
      tkString : Begin
                   SetStrProp(Instance,PropInfo,Value);
                 End;
      tkClass  : Begin
                   O := TObject( GetOrdProp(Instance,PropInfo) );
                   If Assigned(O) And (O Is TStrings) Then
                   Begin
                     TStrings(O).Text := Value;
                   End;
                 End;
    End;
  End;
End;


Function ShiftCount(Mask : UInt64) : UInt8;

Var
  T                       : Integer;
  L                       : Integer;

Begin
  If Mask<>0 Then
  Begin
    T      := 1;
    Result := 0;
    For L := 0 To SizeOf(Mask)*8-1 Do
    Begin
      If (Mask And T)=T Then
        Break;
      T := T Shl 1;
      Inc(Result);
    End;
  End
  Else
  Begin
    Result := 0;
  End;
End;

Function ShiftAndMask(Const Value, Mask : UInt64) : Uint64;

Begin
  If Mask<>0 Then
  Begin
    Result := Value Shl ShiftCount(Mask) And Mask;
  End
  Else
  Begin
    Result := 0;
  End;
End;

Function UnShiftAndMask(Const Value, Mask : UInt64) : Uint64;

Begin
  If Mask<>0 Then
  Begin
    Result := Value And Mask Shr ShiftCount(Mask);
  End
  Else
  Begin
    Result := 0;
  End;
End;

Procedure SetBitField(Var Storage; Bytes : Int8; Const Mask : UInt64; Const Value : UInt64); Overload;

Begin
  Case Bytes Of
    1 : UInt8(Storage)  := (UInt8(Storage)  And Not UInt8(Mask))  or ShiftAndMask(Value,Mask);
    2 : UInt16(Storage) := (UInt16(Storage) And Not UInt16(Mask)) or ShiftAndMask(Value,Mask);
    4 : UInt32(Storage) := (UInt32(Storage) And Not UInt32(Mask)) or ShiftAndMask(Value,Mask);
    8 : UInt64(Storage) := (UInt64(Storage) And Not UInt64(Mask)) or ShiftAndMask(Value,Mask);
  End;
End;

Function GetBitField(Var Storage; Bytes : Int8; Const Mask : UInt64) : Int64; Overload;

Begin
  Case Bytes Of
    1 : Result := UnshiftAndMask(UInt8(Storage),Mask);
    2 : Result := UnshiftAndMask(UInt16(Storage),Mask);
    4 : Result := UnshiftAndMask(UInt32(Storage),Mask);
    8 : Result := UnshiftAndMask(UInt64(Storage),Mask);
  Else
    Result := 0;
  End;
End;

Function ReplaceMarkup(Const Source :String; Const Prefix : String; Values : TStrings) : String;

Var
  Loop                     : Integer;
  StartAt                  : Integer;
  Tag                      : String;
  PrefixLen                : Integer;
  Name                     : String;
  added                    : Boolean;

begin
  PrefixLen := Length(Prefix);

  Result := '';
  Loop := 1;
  While Loop<=Length(Source) Do
  Begin
    StartAt := Loop;
    While (Loop<=Length(Source)) And (Source[Loop]<>'[') Do Inc(Loop);
    If (StartAt<>Loop) Then Result := Result+Copy(Source,StartAt,Loop-StartAt);

    If (Source[Loop]='[') Then
    Begin
      Added := False;
        
      Inc(Loop);
      StartAt := Loop;
      While (Loop<=Length(Source)) And (Source[Loop]<>'[') And (Source[Loop]<>']') Do Inc(Loop);
      Tag := Copy(Source,StartAt,Loop-StartAt);

      If StartsStr(UpperCase(Prefix)+'.',UpperCase(Tag)) Then
      Begin
        Name := Copy(Tag,PrefixLen+2,Length(Tag)-(PrefixLen+1));
        If Values.IndexOfName(Name)<>-1 Then
        Begin
          If Source[Loop]=']' Then Inc(Loop);
          Result := Result+Values.Values[Name];
          Added := True;
        End;
      End;

      If Not Added Then
      Begin
        Result := Result+'['+Tag;
      End;
    End;
  End;
end;

Initialization
  DDUSystem      := TDDUSystem.Create;
  DDUDesktop     := TDDUDesktop.Create;
  DDUEnvironment := TDDUEnvironment.Create;
  DDUEnvironment.LoadFromSystem;
  DDUPath        := TDDUPath.Create;
  DDURTTI        := TDDURTTI.Create;
Finalization
  FreeAndNil(DDURTTI);
  FreeAndNil(DDUPath);
  FreeAndNil(DDUEnvironment);
  FreeAndNil(DDUSystem);
  FreeAndNil(DDUDesktop);
End.


