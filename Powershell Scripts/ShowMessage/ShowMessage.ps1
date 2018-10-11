Function ShowMessage ($messageText)
{
  $objMSGForm = New-Object System.Windows.Forms.Form
  $objMSGForm.Text = "Script Status"
  $objMSGForm.Width = 400
  $objMSGForm.Height = 140
  
  $objMSGLabel1 = New-Object System.Windows.Forms.Label
  $objMSGLabel1.Top = 30
  $objMSGLabel1.Left = 10
  $objMSGLabel1.Size = New-Object System.Drawing.Size(100,110)
  #$objMSGLabel1.Enabled = $false
  $objMSGLabel1.Text = $messageText
  

  $objMSGForm.Controls.AddRange(($objMSGLabel1))
  
  $handler = {$objMSGForm.ActiveControl = $objMSGLabel1}
 
  $objMSGForm.Add_Load($handler)  
  $objMSGForm.ShowDialog()
}

ShowMessage "Test `n`nTest"

