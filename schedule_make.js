const CALENDAR_ID = '';

function myfunction() {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  //date to create
  let dateNum = 26;
  //first year
  let createEventYear = 2024;
  //repeat year
  let repeatyear = 1;

  //event name
  let createEventName = "クレジット引き落とし";
  var subject = 0 


  //-------------------------------main process----------------------------------------------
  //repeat year
  for(let n = 0; n < repeatyear; n++){
    //For12years
    for(let i = 1; i < 1+12; i++){
      //exit
      if (dateNum <= new Date(createEventYear, i, 0).getDate()){
        //day of the week judgement
        let day_of_week = dateweek((createEventYear+n), i, dateNum)

        if (i==1){
          var subject = 12
        }
        else{
          var subject = i-1
        }

        //sunday to fryday
        if (day_of_week>=1 && day_of_week<=5){
          events = calendar.createAllDayEvent((subject + "月分" + createEventName), 
          　　　　　　　　　　　　　　　　　　　　　　　　　　　　　new Date("" + (createEventYear+n) + "/" + i + "/" + dateNum));
        }

        //suturday
        else if (day_of_week==6){
          events = calendar.createAllDayEvent((subject + "月分" + createEventName),
          　　　　　　　　　　　　　　　　　　　　　　　　　　　　　 new Date("" + (createEventYear+n) + "/" + i + "/" + (dateNum+2)));
        }

        //sunday
        else{
          events = calendar.createAllDayEvent((subject + "月分" + createEventName),
          　　　　　　　　　　　　　　　　　　　　　　　　　　　　　 new Date("" + (createEventYear+n) + "/" + i + "/" + (dateNum+1)));
        }

        //argument is minute
        //mail notice
        events.addEmailReminder(60*(6+24));
        //popup notice
        events.addPopupReminder(60*(6+24));
        events.addPopupReminder(60*(6+72));

        //event colar
        events.setColor('11');
        
      }
    }
  }
  
}


//day of the week judgement function
function dateweek(year, month, day) {
  const date = new Date(year, month-1, day);
  const thisDay = date.getDay();
  //const dayArray = ['日', '月', '火', '水', '木', '金', '土'];
  //Logger.log(year + "/" + month + "/" + day + ":" + dayArray[thisDay]);

  return thisDay
}
