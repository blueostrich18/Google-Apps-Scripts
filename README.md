# Google Apps Scripts

## Scripts to help with google sheets

### Script to send email

To add a script to a Google Sheet, Extensions --> Apps Script <br />
![image](https://user-images.githubusercontent.com/54515351/148648359-29d11849-e2b7-4d70-9306-1fd8271a677f.png)

Place the following code in the script window and save it <br />

```javascript
function sendEmailMessage() {
  sender=SpreadsheetApp.getActiveSheet().getRange(2,7).getValue();
  recipient=SpreadsheetApp.getActiveSheet().getRange(2,13).getValue();
  week=SpreadsheetApp.getActiveSheet().getRange(2,6).getValue();
  name=SpreadsheetApp.getActiveSheet().getRange(4,24).getValue();
  var message = {
    to: recipient,
    subject: name+" - Week"+ week + " Complete",
    body: "Week"+week+" Check-In Complete",
    replyTo: sender,
    name: name
  }
  MailApp.sendEmail(message);
}
```

| Cell | Coordinate(row,col) | Variable |
| --- | --- | --- |
| F2 | 2,6 | week |
| G2 | 2,7 | replyTo |
| M2 | 2,13 | recipient |
| X4 | 4,24 | name|

### Create a button
Click Insert-->Drawing  <br />
![image](https://user-images.githubusercontent.com/54515351/148648014-f99b5f9b-e879-4da4-99ed-1a3e19f74d73.png)

Pick a shape and draw it  <br />
![image](https://user-images.githubusercontent.com/54515351/148648064-5a8a4b8d-e659-49e8-93a7-eb438e06c696.png)

Add Text and format the button  <br />
![image](https://user-images.githubusercontent.com/54515351/148648133-09cf481f-77e4-4687-a6fb-981f5020407b.png)

Click "Save and Close"  <br />

Move the new button to where you would like it <br />

Click on the ellipsis(3 dots) of the button, and click "Assign Script" <br />
![image](https://user-images.githubusercontent.com/54515351/148648281-87375ea9-7273-431b-9dcd-7d8bbd543102.png)

Put the name of the script you would like to call when the button is pressed, and click 'OK' <br />
![image](https://user-images.githubusercontent.com/54515351/148648319-010994d9-520e-4af4-a4ba-cf92535421a7.png)




