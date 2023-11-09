"use strict";
const {
  Calendar,
  FocusTrapZone,
  Callout,
  DirectionalHint,
  defaultCalendarStrings,
  DefaultButton,
  ThemeProvider,
  initializeIcons,
  PrimaryButton,
  SharedColors,
} = window.FluentUIReact;
const { useBoolean } = window.FluentUIReactHooks;
// Initialize icons in case this example uses them
initializeIcons();
const CalendarButtonExample = () => {

/*   var userSettings =  Xrm.Utility.getGlobalContext().userSettings; // userSettings is an object with user information.
var current_User_Id = String(userSettings.userId).toLowerCase().replace(/[{}]/g, ''); // The user's unique id
console.log("current user is:",current_User_Id); */

  const [selectedDate, setSelectedDate] = React.useState();
  const [showCalendar, { toggle: toggleShowCalendar, setFalse: hideCalendar }] =
    useBoolean(false);
  const buttonContainerRef = React.useRef(null);
  const onSelectDate = React.useCallback(
    (date, dateRangeArray) => {
      setSelectedDate(date);
      hideCalendar();
    },
    [hideCalendar]
  );
  const onExportClick = () => {
    window.returnValue=selectedDate;
    GetReportExcel(window.returnValue);
    // alert(selectedDate);
  };

  const onCancelClick = () => {
    window.close()
  }

// Return to View React
return React.createElement(
"div",
{ style: {  display: "flex", justifyContent: "center", gap: "2em" } },
React.createElement(
"div",
{ style: { marginTop: "10px", display: "flex", flexDirection: "column", gap: "1.5em" } },
React.createElement(PrimaryButton, { onClick: onExportClick, text: "Export All" }),
React.createElement(DefaultButton, { onClick: onCancelClick, text: "Export Your Profile" }),
React.createElement(DefaultButton, { onClick: onCancelClick, text: "Cancel" })
),
React.createElement(
"div",
{ style: {marginTop: "10px", display: "flex", alignItems: "center", justifyContent: "center", height: "100%", flexDirection: "column"} },
React.createElement(
  "div",
  { ref: buttonContainerRef },
  React.createElement(DefaultButton, {
    onClick: toggleShowCalendar,
    text: !selectedDate
      ? "Choose Month and Year"
      : selectedDate.getMonth()+'/'+selectedDate.getFullYear(),
  })
)
),
showCalendar &&
React.createElement(
  Callout,
  {
    isBeakVisible: false,
    gapSpace: 0,
    doNotLayer: false,
    target: buttonContainerRef,
    directionalHint: DirectionalHint.bottomLeftEdge,
    onDismiss: hideCalendar,
    setInitialFocus: true,
  },
  React.createElement(
    FocusTrapZone,
    { isClickableOutsideFocusTrap: true },
    React.createElement(Calendar, {
      onSelectDate: onSelectDate,
      onDismiss: hideCalendar,
      isMonthPickerVisible: true,
      isDayPickerVisible: false,
      value: selectedDate,
      highlightCurrentMonth: false,
      highlightSelectedMonth: true,
      isDayPickerVisible: false,
      showGoToToday: true,
      // Calendar uses English strings by default. For localized apps, you must override this prop.
      strings: defaultCalendarStrings,
      // -------------------------------
    })
  ),
),
);

};
const CalendarButtonExampleWrapper = () =>
  React.createElement(
    ThemeProvider,
    null,
    React.createElement(CalendarButtonExample, null)
  );
ReactDOM.render(
  React.createElement(CalendarButtonExampleWrapper, null),
  document.getElementById("root")
);
