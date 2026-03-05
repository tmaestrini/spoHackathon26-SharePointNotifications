import { SplitPane, StackV2, TypographyControl } from '@spteck/react-controls-v2';
import { tokens } from '@fluentui/react-components';
import React from "react";

interface ConfigItemProps {
    title: string;
    label?: string;
    children?: React.ReactNode;
}

const ConfigItem: React.FC<ConfigItemProps> = ({ title, label, children }) => {
    return (
        <SplitPane
            direction="horizontal"
            defaultSize="30%"
            resizable={false}
            primaryContent={
                <StackV2 direction="vertical" gap="s"
                    style={{ background: tokens.colorNeutralBackground1, height: '100%' }}>
                    <TypographyControl fontWeight="semibold" fontSize="sm">{title}</TypographyControl>
                    {label && <TypographyControl fontSize="xs" style={{lineHeight: '0.9rem'}} color={tokens.colorNeutralForeground3}>{label}</TypographyControl>}
                </StackV2>
            }
            secondaryContent={
                children ? children : <></>
            }
            styles={{marginBottom: tokens.spacingVerticalL}}
        />
    );
};

export default ConfigItem;