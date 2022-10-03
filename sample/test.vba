Sub Switch_to_Designer_Inputs()
'
' Switch_to_Designer_Inputs Macro
'
    If Worksheets("Designer_Inputs").Visible = False Then
        Worksheets("Designer_Inputs").Visible = True
        Sheets("Designer_Inputs").Select
    End If
    
    If Worksheets("Designer_Inputs").Visible = True Then
        Sheets("Designer_Inputs").Select
    End If

    Worksheets("Salesforce").Visible = False
    Worksheets("AHJ_Review").Visible = False
    
End Sub

Sub Switch_to_AHJ_Review()
' 
'
    If Worksheets("AHJ_Review").Visible = False Then
        Worksheets("AHJ_Review").Visible = True
        Sheets("AHJ_Review").Select
    End If
    
    If Worksheets("AHJ_Review").Visible = True Then
        Sheets("AHJ_Review").Select
    End If

    Worksheets("Designer_Inputs").Visible = False
    Worksheets("Salesforce").Visible = False
    
End Sub

Sub Switch_to_Salesforce_Inputs()
'
'
    If Worksheets("Salesforce").Visible = False Then
        Worksheets("Salesforce").Visible = True
        Sheets("Salesforce").Select
    End If
    
    If Worksheets("Salesforce").Visible = True Then
        Sheets("Salesforce").Select
    End If

    Worksheets("Designer_Inputs").Visible = False
    Worksheets("AHJ_Review").Visible = False
    
End Sub
Sub Clear_Data()
'
' Clears all inital data from this tool
output[ la puta]

    Result = MsgBox("Clear current proposal inputs?", vbYesNo)
    
        If Result = vbYes Then

            ' Clears all inital data from this tool
            
                If Worksheets("Designer_Inputs").Visible = False Then
                    Worksheets("Designer_Inputs").Visible = True
                End If
                
                If Worksheets("Salesforce").Visible = False Then
                    Worksheets("Salesforce").Visible = True
                End If
                
                If Worksheets("AHJ_Review").Visible = False Then
                    Worksheets("AHJ_Review").Visible = True
                End If
                
                If Worksheets("Aurora").Visible = False Then
                    Worksheets("Aurora").Visible = True
                End If
                
                Sheets("Designer_Inputs").Select
                
                Set INPUTS = ThisWorkbook.Sheets("Designer_Inputs")
                
                'Clear Manual Roof Cost data
                Sheets("Designer_Inputs").Range("Roof_Cost_Manual").Value = "0"
                
'                'Reset panel type value
'                Sheets("Designer_Inputs").Range("Panel_Type_Id").Value = "1"
                
'                'Clear Panel Count Value
'                Sheets("Designer_Inputs").Range("Panel_Count").Value = ""
'
'                'Clear Home Owners First Name
'                Sheets("Designer_Inputs").Range("Home_Owner_First_Name").Value = ""
'
'                'Clear Home Owners Last Name
'                Sheets("Designer_Inputs").Range("Home_Owner_Last_Name").Value = ""
'
'                'Clear Utility Bill Saviongs Value
'                Sheets("Designer_Inputs").Range("Utility_Bill_Savings").Value = "0"
'
'                'Clear Incentive data
'                Sheets("Designer_Inputs").Range("Year_1_Production").Value = "0"
                
                'clear the roof sq footage value
                ThisWorkbook.Sheets("Designer_Inputs").Range("Roof_Sq_Footage").Value = ""
                
                'Clear Manual data
                Sheets("Designer_Inputs").Range("Price_Per_Watt_Id").Value = "2"
                
                'Change to the salesforce data input page
                Sheets("Salesforce").Select
                
                'remove the roofer logo check value
                ThisWorkbook.Sheets("Salesforce").Range("Roofer_Logo_Exists").Value = ""
                
                'Clear Opportunity ID
                Sheets("Salesforce").Range("GetObjectId").Value = ""
                
                'Clear Aurora URL
                Sheets("Salesforce").Range("Aurora_URL_1").Value = ""
                
                'Clear Address Info
                Sheets("Salesforce").Range("GetOppty_Name").Value = ""
                
                'Clear Roofer Name
                Sheets("Salesforce").Range("GetAccount_Name").Value = ""
                                
                'Clear Sales Rep Name
                Sheets("Salesforce").Range("GetContact_SalesRep").Value = ""
                
                'Clear Sales Rep Phone Number
                Sheets("Salesforce").Range("GetContact_SalesRepPhone").Value = ""
                
                'Clear Roof Cost
                Sheets("Salesforce").Range("GetOppty_RoofEstimatePrice").Value = ""
                
                'Clear State
                Sheets("Salesforce").Range("GetOppty_State").Value = ""
                
                'Clear Zip
                Sheets("Salesforce").Range("GetOppty_Zip").Value = ""
                
                'Clear OLD AHJ info
                Sheets("Salesforce").Range("AHJVerified").Value = ""
                
                'Update GAF Energy Roofer entry field
                Sheets("Salesforce").Range("GAF_Energy_Roofer_Name_Id").Value = "2"
                
                'Clear previous roofer information
                Sheets("Salesforce").Range("SF_Price_Per_Watt").Value = ""
                Sheets("Salesforce").Range("SF_Approved_For_TS").Value = ""
                Sheets("Salesforce").Range("SF_Bulk_Partner").Value = ""
                Sheets("Salesforce").Range("SF_Product_Plus_Design").Value = ""
                Sheets("Salesforce").Range("SF_Design_Template_Message").Value = ""
                Sheets("Salesforce").Range("SF_Design_Notes").Value = ""
                Sheets("Salesforce").Range("SF_Design_Adjustments").Value = ""
                Sheets("Salesforce").Range("SF_Roofer_Specific_Guidelines").Value = ""
                Sheets("Salesforce").Range("SF_Proposal_SKUs").Value = ""
                
                'Change to the AHJ Review page
                Sheets("AHJ_Review").Select
                
                'Delete current inputs
                Sheets("AHJ_Review").Range("Install_Address_MANUAL").Value = ""
                Sheets("AHJ_Review").Range("Install_State_MANUAL").Value = ""
                
                'Clear OLD Aurora info
                Sheets("Aurora").Range("Aurora_First_Name").Value = ""
                Sheets("Aurora").Range("Aurora_Second_Name").Value = ""
                Sheets("Aurora").Range("Aurora_Lifetime_Savings").Value = ""
                Sheets("Aurora").Range("Aurora_Year_1_Production").Value = ""
                Sheets("Aurora").Range("Aurora_System_Size").Value = ""
                Sheets("Aurora").Range("Aurora_Roof_Area").Value = ""
                Sheets("Aurora").Range("Flat_System_Cost").Value = ""
                Sheets("Aurora").Range("Panel_Name").Value = ""

                Worksheets("AHJ_Review").Visible = False
                Worksheets("Designer_Inputs").Visible = False
                Worksheets("Aurora").Visible = False
                Worksheets("Salesforce").Visible = True

                
                Range("A1").Select
        
        End If

End Sub
