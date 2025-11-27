function listCalendars() {
  var calendars = CalendarApp.getAllCalendars();
  console.log('Found ' + calendars.length + ' calendars');
  for (var i = 0; i < calendars.length; i++) {
    console.log('ID: ' + calendars[i].getId() + ', Name: ' + calendars[i].getName());
  }
}
