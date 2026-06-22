; AKSA 10X FASTER InnoSetup Installer
#define MyAppName "AKSA 10X FASTER"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "AKSA DIGITAL POINT"
#define MyAppURL "https://aksa10xfaster.github.io"

[Setup]
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={localappdata}\Aksa10xFaster
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=.
OutputBaseFilename=AKSA_10X_FASTER_Setup
SetupIconFile=app.ico
WizardImageFile=logo.png
WizardSmallImageFile=icon.png
UninstallDisplayIcon={app}\bin\WordAddIn.dll
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
DisableWelcomePage=no
LicenseFile=license.txt

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "bin\WordAddIn.dll"; DestDir: "{app}\bin"; Flags: ignoreversion
Source: "icons\*.png"; DestDir: "{app}\icons"; Flags: ignoreversion
Source: "icons\*.png"; DestDir: "{userappdata}\Aksa10xFaster\Icons"; Flags: ignoreversion
Source: "Document Template\*"; DestDir: "{app}\Document Template"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "Document Template\*"; DestDir: "{userappdata}\Aksa10xFaster\Document Template"; Flags: ignoreversion recursesubdirs createallsubdirs

[Dirs]
Name: "{userappdata}\Aksa10xFaster\Icons"
Name: "{userappdata}\Aksa10xFaster\Document Template"

[Registry]
; Copy icons to APPDATA
Root: HKCU; Subkey: "Software\Aksa10xFaster"; ValueType: none; Flags: createvalueifdoesntexist

; COM CLSID registration
Root: HKCU; Subkey: "Software\Classes\CLSID\{{33F39564-802B-412B-A713-E947F98DCBB8}}"; ValueType: string; ValueData: "Aksa10xFaster.Connect"; Flags: createvalueifdoesntexist
Root: HKCU; Subkey: "Software\Classes\CLSID\{{33F39564-802B-412B-A713-E947F98DCBB8}}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"; Flags: createvalueifdoesntexist
Root: HKCU; Subkey: "Software\Classes\CLSID\{{33F39564-802B-412B-A713-E947F98DCBB8}}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{33F39564-802B-412B-A713-E947F98DCBB8}}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "WordAddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{33F39564-802B-412B-A713-E947F98DCBB8}}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "WordToExeAddin.WordAddIn"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{33F39564-802B-412B-A713-E947F98DCBB8}}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{33F39564-802B-412B-A713-E947F98DCBB8}}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "file:///{app}\bin\WordAddIn.dll"

; ProgId registration
Root: HKCU; Subkey: "Software\Classes\Aksa10xFaster.Connect"; ValueType: string; ValueData: "AKSA 10X FASTER"; Flags: createvalueifdoesntexist
Root: HKCU; Subkey: "Software\Classes\Aksa10xFaster.Connect\CLSID"; ValueType: string; ValueData: "{{33F39564-802B-412B-A713-E947F98DCBB8}}"

; Word Add-in registration
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\Aksa10xFaster.Connect"; ValueType: string; ValueName: "Description"; ValueData: "AKSA 10X FASTER - Word shortcut tools"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\Aksa10xFaster.Connect"; ValueType: string; ValueName: "FriendlyName"; ValueData: "AKSA 10X FASTER"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\Aksa10xFaster.Connect"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\Aksa10xFaster.Connect"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0"

; Word 2016+ specific
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Word\Addins\Aksa10xFaster.Connect"; ValueType: string; ValueName: "Description"; ValueData: "AKSA 10X FASTER - Word shortcut tools"
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Word\Addins\Aksa10xFaster.Connect"; ValueType: string; ValueName: "FriendlyName"; ValueData: "AKSA 10X FASTER"
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Word\Addins\Aksa10xFaster.Connect"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Word\Addins\Aksa10xFaster.Connect"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0"

; No [Run] or [UninstallRun] sections — file copying done in [Files] to avoid cmd.exe/xcopy (AV triggers)

[Code]
var
  ActivationPage: TInputOptionWizardPage;
  LicenseKeyPage: TInputQueryWizardPage;

procedure InitializeWizard;
begin
  // Custom page 1: Trial or License Key
  ActivationPage := CreateInputOptionPage(
    wpLicense,  // After license agreement page
    'Activation',
    'Choose how to activate AKSA 10X FASTER',
    'Select an option below:',
    True,  // Radio buttons (not checkboxes)
    False  // Not with option values
  );
  ActivationPage.Add('30 Days Free Trial'#13#13'Full features for 30 days. No credit card required. After trial, JPG save, Math tools, and AI features will be locked.');
  ActivationPage.Add('Activate License Key'#13#13'Enter your 25-character license key to unlock all features permanently.');
  ActivationPage.SelectedValueIndex := 0;

  // Custom page 2: License Key input (only shown if "Activate License Key" selected)
  LicenseKeyPage := CreateInputQueryPage(
    ActivationPage.ID,
    'License Key',
    'Enter your license key',
    'Please enter your 25-character license key (e.g. XXXXX-XXXXX-XXXXX-XXXXX-XXXXX):'
  );
  LicenseKeyPage.Add('License Key:', False);
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  // Skip license key page if trial selected
  if PageID = LicenseKeyPage.ID then
  begin
    Result := ActivationPage.SelectedValueIndex <> 1;
  end
  else
    Result := False;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  Key: String;
  Clean: String;
  I: Integer;
  C: Char;
  Valid: Boolean;
begin
  Result := True;

  if CurPageID = LicenseKeyPage.ID then
  begin
    Key := Trim(LicenseKeyPage.Values[0]);
    if Key = '' then
    begin
      MsgBox('Please enter a license key, or go back and select "30 Days Free Trial".', mbError, MB_OK);
      Result := False;
      Exit;
    end;

    Clean := '';
    for I := 1 to Length(Key) do
    begin
      C := Key[I];
      if (C <> '-') and (C <> ' ') then
        Clean := Clean + C;
    end;
    Clean := UpperCase(Clean);

    if Length(Clean) <> 25 then
    begin
      MsgBox('License key must be 25 characters (A-Z, 0-9).', mbError, MB_OK);
      Result := False;
      Exit;
    end;

    Valid := True;
    for I := 1 to 25 do
    begin
      C := Clean[I];
      if not ((C >= 'A') and (C <= 'Z') or (C >= '0') and (C <= '9')) then
      begin
        Valid := False;
        Break;
      end;
    end;

    if not Valid then
    begin
      MsgBox('License key contains invalid characters. Use only A-Z and 0-9.', mbError, MB_OK);
      Result := False;
      Exit;
    end;

    LicenseKeyPage.Values[0] := Clean;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  Key: String;
begin
  if CurStep = ssPostInstall then
  begin
    // Fix ProgId\CLSID value (InnoSetup constant expansion strips braces)
    RegWriteStringValue(HKCU, 'Software\Classes\Aksa10xFaster.Connect\CLSID', '', '{33F39564-802B-412B-A713-E947F98DCBB8}');

    // Save license key if provided
    if ActivationPage.SelectedValueIndex = 1 then
    begin
      Key := LicenseKeyPage.Values[0];
      if Key <> '' then
      begin
        RegWriteStringValue(HKCU, 'Software\Aksa10xFaster', 'LicenseKey', Key);
      end;
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // Clean up AppData folders
    DelTree(ExpandConstant('{userappdata}\Aksa10xFaster\Icons'), True, True, True);
    DelTree(ExpandConstant('{userappdata}\Aksa10xFaster\Document Template'), True, True, True);

    // Clean up registry
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\CLSID\{{33F39564-802B-412B-A713-E947F98DCBB8}}');
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\Aksa10xFaster.Connect');
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Microsoft\Office\Word\Addins\Aksa10xFaster.Connect');
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Microsoft\Office\16.0\Word\Addins\Aksa10xFaster.Connect');
  end;
end;
