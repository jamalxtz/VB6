Public Function CalculaDV(chave43 As String)
    Dim indice, multiplicador As Integer    
	Dim soma, resto, digito_verificador As Integer
	'Zera a soma    
	soma = 0    
	'Multiplicador inicia com 9    
	multiplicador = 2        
	'Multiplica do 43° até o 1° caractere da chave    
	For indice = Len(chave43) To 1 Step-1        
		'Multiplica cada digito da chave pelo multiplicador correspondente e soma        
		soma = soma + (Mid(chave43, indice, 1) * multiplicador)        
		multiplicador = multiplicador + 1         
		'Se multiplicador chegou a 2, volta para 9        
		If(multiplicador > 9) Then multiplicador = 2    
	Next indice     
	'Pega o resto da divisão através da função mod    
	resto = soma Mod 11        
	'Dígito verificador é o resultado da subtração 11 - resto    
	digito_verificador = 11 - resto        
	'Testa se o DV é maior = 10    
	If(digito_verificador >= 10) Then digito_verificador = 0        
	'Retorna o DV    
	CalculaDV = Abs(digito_verificador)
End Function


'------------- https://www.onlinegdb.com/online_vb_compiler --------------------------------------------------------------------------

Module VBModule
    Sub Main()
        dim chave43 As String
        chave43 = "5221033366126200015665001000479071943363541"
        Dim indice, multiplicador As Integer    
    	Dim soma, resto, digito_verificador As Integer
    	'Zera a soma    
    	soma = 0    
    	'Multiplicador inicia com 9    
    	multiplicador = 2        
    	'Multiplica do 43° até o 1° caractere da chave    
    	For indice = Len(chave43) To 1 Step-1        
    		'Multiplica cada digito da chave pelo multiplicador correspondente e soma        
    		soma = soma + (Mid(chave43, indice, 1) * multiplicador)        
    		multiplicador = multiplicador + 1         
    		'Se multiplicador chegou a 2, volta para 9        
    		If(multiplicador > 9) Then multiplicador = 2    
    	Next indice     
    	'Pega o resto da divisão através da função mod    
    	resto = soma Mod 11        
    	'Dígito verificador é o resultado da subtração 11 - resto    
    	digito_verificador = 11 - resto        
    	'Testa se o DV é maior = 10    
    	If(digito_verificador >= 10) Then digito_verificador = 0        
    	'Retorna o DV    
    	Console.WriteLine("Código verificador: " & digito_verificador)
    End Sub
End Module