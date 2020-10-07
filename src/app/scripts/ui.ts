/* eslint-disable @typescript-eslint/no-explicit-any */

export type { ComponentSlotStyle, ThemePrepared } from "@fluentui/styles";
export type { ComponentEventHandler, ShorthandValue, ShorthandCollection } from "@fluentui/react-northstar/dist/es/types";
export type { AlertProps } from "@fluentui/react-northstar/dist/es/components/Alert/Alert";
export type { DropdownProps } from "@fluentui/react-northstar/dist/es/components/Dropdown/Dropdown";
export type { DropdownItemProps } from "@fluentui/react-northstar/dist/es/components/Dropdown/DropdownItem";
export type { InputProps } from "@fluentui/react-northstar/dist/es/components/Input/Input";
export type { RadioGroupItemProps } from "@fluentui/react-northstar/dist/es/components/RadioGroup/RadioGroupItem";
export type { TableRowProps } from "@fluentui/react-northstar/dist/es/components/Table/TableRow";
export type { TableCellProps } from "@fluentui/react-northstar/dist/es/components/Table/TableCell";

export { Alert } from "@fluentui/react-northstar/dist/es/components/Alert/Alert";
export { Button } from "@fluentui/react-northstar/dist/es/components/Button/Button";
export { Card } from "@fluentui/react-northstar/dist/es/components/Card/Card";
export { CardHeader } from "@fluentui/react-northstar/dist/es/components/Card/CardHeader";
export { CardFooter } from "@fluentui/react-northstar/dist/es/components/Card/CardFooter";
export { CardBody } from "@fluentui/react-northstar/dist/es/components/Card/CardBody";
export { Checkbox } from "@fluentui/react-northstar/dist/es/components/Checkbox/Checkbox";
export { Dialog } from "@fluentui/react-northstar/dist/es/components/Dialog/Dialog";
export { Divider } from "@fluentui/react-northstar/dist/es/components/Divider/Divider";
export { Dropdown } from "@fluentui/react-northstar/dist/es/components/Dropdown/Dropdown";
export { Flex } from '@fluentui/react-northstar/dist/es/components/Flex/Flex';
export { Header } from "@fluentui/react-northstar/dist/es/components/Header/Header";
export { Input } from "@fluentui/react-northstar/dist/es/components/Input/Input";
export { Loader } from "@fluentui/react-northstar/dist/es/components/Loader/Loader";
export { Provider } from "@fluentui/react-northstar/dist/es/components/Provider/Provider";
export { RadioGroup } from "@fluentui/react-northstar/dist/es/components/RadioGroup/RadioGroup";
export { Segment } from "@fluentui/react-northstar/dist/es/components/Segment/Segment";
export { Table } from "@fluentui/react-northstar/dist/es/components/Table/Table";
export { Text } from "@fluentui/react-northstar/dist/es/components/Text/Text";
export { TriangleDownIcon } from '@fluentui/react-icons-northstar/dist/es/components/TriangleDownIcon';
export { TriangleUpIcon } from '@fluentui/react-icons-northstar/dist/es/components/TriangleUpIcon';
export { AcceptIcon } from '@fluentui/react-icons-northstar/dist/es/components/AcceptIcon';
export { BanIcon } from '@fluentui/react-icons-northstar/dist/es/components/BanIcon';


export { DatePicker } from "office-ui-fabric-react/lib/components/DatePicker";
export { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/components/FocusZone';
export { Link } from "office-ui-fabric-react/lib/components/Link";
export { List } from "office-ui-fabric-react/lib/components/List";
export type { IList } from "office-ui-fabric-react/lib/components/List";

import { teamsDarkTheme as darkTheme } from "@fluentui/react-northstar/dist/es/themes/teams-dark";
import { teamsTheme as lightTheme } from "@fluentui/react-northstar/dist/es/themes/teams";
import { teamsHighContrastTheme as highContrastTheme } from "@fluentui/react-northstar/dist/es/themes/teams-high-contrast";
import { mergeThemes } from "@fluentui/styles/dist/es/mergeThemes";
export { mergeStyles } from "@uifabric/merge-styles/lib/mergeStyles";
import type { ThemeInput, ComponentVariablesPrepared } from "@fluentui/styles/dist/es/types";

import { PageStyles } from './Page';
export { Page } from './Page';


const CardVariables: ComponentVariablesPrepared = (siteVariables) => ({
    backgroundColor: siteVariables?.colorScheme.default.background,
    backgroundColorHover: siteVariables?.colorScheme.default.background,
    margin: "0 0 16px 0",
    borderStyle: 'solid',
    borderWidth: '2px 0 0 0',
    boxShadow: '0 1px 1px 1px rgba(34,36,38,.15)',
});


const lightCustomizations: ThemeInput = {
    componentVariables: {
        Card: CardVariables,
    },
    componentStyles: {
        Page: PageStyles,
    },
};

const darkCustomizations = {
    componentVariables: {
        Card: CardVariables,
    },
    componentStyles: {
        Page: PageStyles,
    },
};

const highContractCustomizations = {
    componentVariables: {
        Card: CardVariables,
    },
    componentStyles: {
        Page: PageStyles,
    },
};

export const teamsTheme = mergeThemes(lightTheme, lightCustomizations);
export const teamsDarkTheme = mergeThemes(darkTheme, darkCustomizations);
export const teamsHighContrastTheme = mergeThemes(highContrastTheme, highContractCustomizations);

