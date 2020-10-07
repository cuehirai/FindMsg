/* eslint-disable react/prop-types */
/* eslint-disable react/display-name */
import * as React from "react";
import { createComponent } from '@fluentui/react-northstar/dist/es/utils/createComponent';
import { Flex } from './ui';

const PageStyles = {
    // eslint-disable-next-line @typescript-eslint/explicit-module-boundary-types
    root: ({ theme: { siteVariables } }) => ({
        backgroundColor: siteVariables.colorScheme.default.background2,
        minHeight: "100vh",
    }),
};

const Page = createComponent({
    displayName: 'Page',
    render: ({ config, children }) => {
        const { classes } = config
        return <Flex gap="gap.medium" fill column padding="padding.medium" className={classes.root}>{children}</Flex>
    },
});

export { Page, PageStyles };