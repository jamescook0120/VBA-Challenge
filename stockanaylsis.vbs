Homework(

dim volume as long
dim price as double
dim tickerpos as interger
dim priceopen as double
dim priceend as double

    volume = 0
    tickerpos = 2


for i = 2 to rows.count
    priceopen = cells(i,3).value

    
    if cells(i,1).value = cells(i+1).value Then
        volume = volume + cells(i,7).value
        priceopen = cells(i,3).value

    ElseIf cells(i,1).value <> cells(i+1,1).value Then
        cells(tickerpos,9).value = cells(i,1).Value
        
        priceclose = cells(i,6).value

        cells(tickerpos,10).value = priceopen - priceclose
        cells(tickerpos,11).value = priceopen/priceclose
        cells(tickerpos,12.value) = volume
        tickerpos = tickerpos + 1

        volume = 0 

next i 

)