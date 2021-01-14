dim i,tipo,n
call inicio()
sub inicio() 
tipo=inputbox("Tecle (F) For ou (W) while","ATENÇÂO")
select case tipo
        case "f","F":
		msgbox("Laço while")
		for i=1 to 10 step 1
		randomize(second(time))
		n=int(rnd*100) +1
		msgbox(n)
		next
		msgbox(i)
		case "w","W":
		i=1
		do while i<10
		msgbox(i)
		i=i+1	 
		loop
case else
msgbox("opcao invalida")		
call inicio
end select
end sub

