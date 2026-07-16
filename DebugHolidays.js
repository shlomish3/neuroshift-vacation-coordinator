function testGetVacations() {
  const res = getExistingVacations();
  Logger.log("Holidays: " + JSON.stringify(res.holidays));
}
