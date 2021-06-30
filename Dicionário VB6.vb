=====================
Dicionario VB6


http://www.macoratti.net/d150902.htm

Você sabia que o VB possui um objeto Dicionário ( Dictionary ) ? Pois é , tem sim. Vamos dar uma olhada nele...

O objeto dictionary é um componente da : Microsoft Scripting Library , e , para poder usá-lo no seu projeto você terá que referenciar esta livraria . (SCRRUN.DLL).

O objeto Dicionário(Dictionary) é semelhante ao objeto Colletion em funcionalidade e propósitos. O Dictionary porém ofere algumas funcionalidades que não estão disponíveis no objeto Colletion. Dentre elas podemos citar:

A opção de especificar um método de comparação para chaves (Keys).(Case sensitive)
Um método para determinar se um objeto existe no Dicionário.
Um método para extrair todas as chaves em um array (vetor)
Um método para extrair todos os items em um array.
Um método para alterar o valor de uma chave.
Um método para remover todos os items de um Dicionário.
As chaves do objeto Dictionary não são limitados ao tipo de dados String.
Obs : Se você usar a propriedaede Item em um Dicionário para referenciar um chave que não existe , a chave será incluida no Dicionário. Se fizer a mesma coisa em uma Coleção vai obter um erro.

Nota: O VB5 NÃO vem com a Microsoft Scripting library , você vai ter que instalar fazendo o download do site da Microsoft.

Usando o objeto Dictionary

Inicie um novo projeto no VB
Inclua uma referência a Microsoft Scripting Runtime.
Inclua um módulo padrão ao seu projeto
No menu Project , selecione Project1.Properties e altera o objeto Startup para Sub Main.
- Na seção General Declarations do formulário insira o código que define o objeto dicionário

Option Explicit
Dim dicionario As Dictionary

- A seguir em Sub Main inclua o código que realiza algumas operações com o objeto Dictionary:

Sub Main()
Dim keyArray, itemArray, elemento

Set dicionario = New Dictionary
With dicionario
   'define o modo de comparação
   .CompareMode = BinaryCompare
   'inclui um tem com argumentos nomeados
   .Add Key:="macoratti", Item:=22
   'inclui um item sem argumentos nomeados
   .Add "miriam", 33	

   'Verificando case sensitivity e o método method
   'macoratti existe ?
   Debug.Print "MACORATTI existe ? -> " & .Exists("macoratti")
   'alterando o valor da chave
   .Key("macoratti") = "Jefferson Andre"
   'Jefferson Andre existe?
   Debug.Print "Jefferson Andre existe ? -> " & .Exists("Jefferson Andre")
  
   'extrai as chaves em um vetor
   Debug.Print "Vetor de chaves"
   keyArray = .Keys
   For Each elemento In keyArray
      Debug.Print elemento
   Next

   'extrai itens do vetor
   Debug.Print "Vetor de itens"
   itemArray = .Items
   For Each elemento In itemArray
      Debug.Print elemento
   Next

   'limpa o dicionario
   .RemoveAll
   Debug.Print dicionario.Count & " Itens no dicionario"

End With
Set dicionario = Nothing
End Sub
Ao executar o projeto teremos o seguinte resultado na janela de depuração :

MACORATTI existe ? -> True
Jefferson Andre existe ? -> True
Vetor de chaves
Jefferson Andre
miriam
Vetor de itens
22
33
0 Itens no dicionario
Só isto... Até mais... 