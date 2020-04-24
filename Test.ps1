$ol = New-Object -comObject Outlook.Application
$mail = $ol.CreateItem(0)
$mail.Subject = "Top demand apps-SOURCE CLARIFICATION"
$mail.HTMLBody="<html><head></head><body><b>Joseph</b></body></Html>"
$mail.save()

$inspector = $mail.GetInspector
$inspector.Display()
