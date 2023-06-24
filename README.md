# Module_Two_Challenge
# I had a one hour session with one of the program's tutors. She showed me how I could use a "For Each" loop to make sure my script looped through all the spreadsheets. (For Each CS In ActiveWorkbook.Worksheets)

#She also helped me devise code to Calculate Yearly Change, Percent Change, Greatest Increase and Decrease, as seen below:
#Close_Price = CS.Cells(i, 6).Value
            
            Yearly_Change = Close_Price - Open_Price
            
            If Open_Price <> 0 Then
            
                Percent_Change = (Yearly_Change / Open_Price) * 100
                
                Else
                    Percent_Change = 0
                    
                End If
            
            If Percent_Change > Greatest_Increase Then
            
                Greatest_Increase = Percent_Change
                
                Greatest_Increase_Ticker = Ticker
                
                End If
                
            If Percent_Change < Greatest_Decrease Then
            
                Greatest_Decrease = Percent_Change
                
                Greatest_Decrease_Ticker = Ticker
                
                End If

#I also had a 15 minute session with a BCS learning assistant who introduced me to the "Select Case" method, which I used in my conditional formatting:

Select Case Yearly_Change
            
            Case Is > 0
            
            CS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            
             CS.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            
            Case Is < 0
            
            CS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
            CS.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            
            End Select
            

