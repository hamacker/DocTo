unit utils_docto;

interface

// Function to use DocTo in Pascal/Delphi projects, returning Empty string(Sucess)
//   or a String Message Error(if Fails), parameters are the same in cmd version, ex:
//   bOk:=Docto('-f C:\Directory\MyFile.doc -O "C:\Output Directory\MyTextFile.txt" -T wdFormatText');
//   By Gladiston Santana sirhamacker at gmail dot com
//
function DocTo(ACmdParam:String):String;

implementation
uses
  SysUtils,
  Classes,
  ActiveX,
  WordUtils,
  MainUtils,
  ResourceUtils,
  PathUtils,
  datamodSSL,
  ExcelUtils,
  Word_TLB_Constants,
  Excel_TLB_Constants;

function DocTo(ACmdParam:String):String;
var
  i, Converter : integer;
  paramlist : TStringlist;
  DocConv : TWordDocConverter;
  XLSConv : TExcelXLSConverter;
  LogResult : String;
  L:TStringList;
  S:String;
begin
  Result:='';
  L:=TStringList.Create;
  L.Text:=StringReplace(ACmdParam, #32, sLineBreak, [rfReplaceAll]);

  paramlist := TStringlist.create;

  try
   try
     DocConv := TWordDocConverter.Create;
     XLSConv := TExcelXLSConverter.Create;
    try

      for i := 0 to Pred(L.Count) do
      begin
        S:=Trim(L[i]);
        if S<>'' then        
          paramlist.Add(S);
      end;

      CoInitialize(nil);

      Converter := DocConv.ChooseConverter(ParamList);

      if Converter = MSWord then
      begin
        DocConv.LoadConfig(paramlist);
        LogResult :=  DocConv.Execute;
        DocConv.log( LogResult );
      end
      else begin
        XLSConv.LoadConfig(ParamList);
        LogResult :=  XLSConv.Execute;
        XLSConv.log( LogResult );
      end;

      CoUninitialize;
    finally
      DocConv.free;
      XLSConv.Free;
    end;

   except on E: Exception do
    begin
      Result:=E.ClassName+':'+E.Message;
    end;

   end;
  finally
    L.Free;
    paramlist.Free;
  end;
end;

end.
