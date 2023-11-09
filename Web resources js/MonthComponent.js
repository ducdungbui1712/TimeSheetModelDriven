"use strict";
const {
  Calendar,
  DateRangeType,
  defaultCalendarStrings,
  ThemeProvider,
  initializeIcons,
  PrimaryButton,
} = window.FluentUIReact;

initializeIcons();

const CalendarInlineMonthOnlyExample = () => {
  const [selectedDateRange, setSelectedDateRange] = React.useState();
  const [selectedDate, setSelectedDate] = React.useState();
  const onSelectDate = React.useCallback((date, dateRangeArray) => {
    setSelectedDate(date);
    setSelectedDateRange(dateRangeArray);
  }, []);

  let dateRangeString = "Not set";
  if (selectedDateRange) {
    const rangeStart = selectedDateRange[0];
    const rangeEnd = selectedDateRange[selectedDateRange.length - 1];
    dateRangeString =
      rangeStart.toLocaleDateString() + "-" + rangeEnd.toLocaleDateString();
  }

  const onExportClick = () => {
    alert(selectedDate);
  };

  return React.createElement(
    "div",
    { style: { height: "auto", display: "flex", justifyContent: "center", alignItems: "center", flexDirection: "column" } },
     React.createElement(
      "div",
      null,
      "Selected date: ",
      (selectedDate === null || selectedDate === void 0
        ? void 0
        : selectedDate.toLocaleString()) || "Not set"
    ),
    React.createElement("div", null, "Selected range: ", dateRangeString),
    React.createElement(Calendar,
      {
      dateRangeType: DateRangeType.Month,
      showGoToToday: true,
      highlightSelectedMonth: true,
      isDayPickerVisible: false,
      onSelectDate: onSelectDate,
      value: selectedDate,
      strings: defaultCalendarStrings,
      highlightSelectedMonth: true,
    
    }),
    React.createElement("div", { style: { marginTop: "10px", display: "flex", justifyContent: "center", width: "100%" } },
      React.createElement(PrimaryButton, { onClick: onExportClick, text: "Export" })
    )
  );
};

const CalendarInlineMonthOnlyExampleWrapper = () =>
  React.createElement(
    ThemeProvider,
    null,
    React.createElement(CalendarInlineMonthOnlyExample, null)
  );

ReactDOM.render(
  React.createElement(CalendarInlineMonthOnlyExampleWrapper, null),
  document.getElementById("root")
);
