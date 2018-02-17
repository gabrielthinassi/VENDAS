unit Exemploes;

interface

implementation


//1 - Controle de Transa��o
{
var
  Trans: TDBXTransaction;
begin
  Trans := SQLConnection1.BeginTransaction;
  try
    with CDS1 do
    begin
      Append;
      ...
      Post;
    end;

    If CDS1.ApplyUpdates(0) = 0 then
      SQLConnection1.CommitFreeAndNil(Trans);
      ShowMessage('Os Dados foram gravados com sucesso!');
  Finally
    SQLConnection1.RollbackFreeAndNil(Trans);
  end;
end;
}

// CTRL + F7
// Permite visualizar o retorno de uma Fun��o em Debug Mode;




// 2 - Procedures e Functions
// https://www.devmedia.com.br/conceitos-basicos-sobre-procedures-e-functions-no-delphi/25303



end.
