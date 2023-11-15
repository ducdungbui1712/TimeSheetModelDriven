"use strict";

const {
  Calendar,
  FocusTrapZone,
  mergeStyleSets,
  Dropdown,
  DropdownMenuItemType,
  ThemeProvider,
  DefaultButton,
  PrimaryButton,
  initializeIcons,
  FontSizes,
  TextField,
  ITextFieldStyleProps,
  ITextFieldStyles,
  createTheme,
  Callout,
  DirectionalHint,
  defaultCalendarStrings

} = window.FluentUIReact;

// Initialize icons in case this example uses them
initializeIcons();
const { useBoolean } = window.FluentUIReactHooks;
// Styles for the combined form
const styles = mergeStyleSets({
  root: { selectors: { '> *': { marginBottom: 15 } } },
  control: { width: 300, marginBottom: 10 },
  dropdown: { width: 300 },
});

// Dropdown options
const dropdownStyles = { dropdown: styles.dropdown };
const dropdownControlledOptions = [
  { key: 'exportTypes', text: 'Export Types', itemType: DropdownMenuItemType.Header },
  { key: 'exportAll', text: 'All' },
  { key: 'export', text: 'Yourself' }
];


// Combined form component
const CombinedForm = () => {

  const [selectedDropdownItem, setSelectedDropdownItem] = React.useState();

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

  const onDropdownChange = (event, item) => {
    setSelectedDropdownItem(item);
    console.log(item.key)
  };

  // const onClearClick = () => {
  //   setSelectedDate(null);
  //   setSelectedDropdownItem({ key: null });
  // };
  
  const onExportClick = () => {
    // Logic for exporting data
    let nameButton = selectedDropdownItem.key
    window.returnValue=selectedDate;
    GetReportExcel(window.returnValue,nameButton);

  };
  
  const onCancelClick = () => {
    // Logic for canceling
    console.log("Cancel clicked");
    window.close()
  };

  const theme = createTheme({
    palette: {
      themePrimary: '#0078d4',
      neutralPrimary: '#323130',
      neutralLighterAlt: '#faf9f8',
    },
  });
  
  const inputStyles = {
   
    border: `1px solid ${theme.palette.neutralPrimary}`,
    boxSizing: 'border-box',
    borderRadius: '2px',
    paddingLeft: '12px',
    fontSize: '14px',
    marginBottom: '2em',
    backgroundColor: theme.palette.neutralLighterAlt,
  };
  

  return React.createElement(
    "div",
    { className: styles.root },
    //New DatePicker
    React.createElement(
      "div",
      { className: styles.control, ref: buttonContainerRef },
      React.createElement(DefaultButton, {
        onClick: toggleShowCalendar,
        text: !selectedDate
          ? "Choose Month and Year"
          : selectedDate.getMonth()+1+'/'+selectedDate.getFullYear(),
      })
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


    // Dropdown
    React.createElement(
      Dropdown,
      {
        label: "Select export types",
        selectedKey: selectedDropdownItem ? selectedDropdownItem.key : undefined,
        onChange: onDropdownChange,
        placeholder: "Select an option",
        options: dropdownControlledOptions,
        styles: dropdownStyles,
      }
    ),

    // Nút Clear
    // React.createElement(
    //   DefaultButton,
    //   { text: "Clear", onClick: onClearClick, style: { marginBottom: "2em", display: "block" } }
    // ),
    React.createElement(
      "div",
      { style: { display: "flex", gap: "2em"} },
      // Nút Export
      React.createElement(
        PrimaryButton,
        { text: "Export", onClick: onExportClick }
      ),
      // Nút Cancel
      React.createElement(
        DefaultButton,
        { text: "Cancel", onClick: onCancelClick }
      )
    )
  );
};

// Wrapper component for the combined form
const CombinedFormWrapper = () =>
  React.createElement(
    ThemeProvider,
    null,
    React.createElement(CombinedForm, null)
  );

// Render the combined form into the HTML element with the id 'root'
ReactDOM.render(
  React.createElement(CombinedFormWrapper, null),
  document.getElementById('root')
);
