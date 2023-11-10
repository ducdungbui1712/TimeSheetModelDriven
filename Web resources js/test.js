"use strict";

const {
  DatePicker,
  mergeStyleSets,
  defaultDatePickerStrings,
  Dropdown,
  DropdownMenuItemType,
  ThemeProvider,
  DefaultButton,
  PrimaryButton,
  initializeIcons,
} = window.FluentUIReact;

// Initialize icons in case this example uses them
initializeIcons();

// Styles for the combined form
const styles = mergeStyleSets({
  root: { selectors: { '> *': { marginBottom: 15 } } },
  control: { maxWidth: 300, marginBottom: 15 },
  dropdown: { width: 300 },
});

// Date picker options
const datePickerOptions = {
  label: "Select a month",
  allowTextInput: true,
  ariaLabel: "Select a date",
  className: styles.control,
  strings: defaultDatePickerStrings
};

// Dropdown options
const dropdownStyles = { dropdown: styles.dropdown };
const dropdownControlledOptions = [
  { key: 'exportTypes', text: 'Export Types', itemType: DropdownMenuItemType.Header },
  { key: 'all', text: 'All' },
  { key: 'yourself', text: 'Yourself' }
];


// Combined form component
const CombinedForm = () => {
  const [datePickerValue, setDatePickerValue] = React.useState();
  const [selectedDropdownItem, setSelectedDropdownItem] = React.useState();

  const onDatePickerSelect = (selectedDate) => {
    setDatePickerValue(selectedDate);
  };

  const onDropdownChange = (event, item) => {
    setSelectedDropdownItem(item);
  };

  const onClearClick = () => {
    setDatePickerValue(null);
    setSelectedDropdownItem({ key: null });
  };
  
  const onExportClick = () => {
    // Logic for exporting data
    console.log("Exporting data:", datePickerValue, selectedDropdownItem);
  };
  
  const onCancelClick = () => {
    // Logic for canceling
    console.log("Cancel clicked");
  };

  return React.createElement(
    "div",
    { className: styles.root },
    // Date Picker
    React.createElement(
      DatePicker,
      Object.assign({}, datePickerOptions, {
        value: datePickerValue,
        onSelectDate: onDatePickerSelect
      })
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
    React.createElement(
      DefaultButton,
      { text: "Clear", onClick: onClearClick, style: { marginBottom: "2em" } }
    ),
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
