
import * as React from 'react';
import { Info24Regular, Dismiss16Regular, ErrorCircleRegular, ErrorCircle24Regular, CheckmarkCircle24Regular, Dismiss24Regular, DismissCircle24Regular } from '@fluentui/react-icons';
import { StackV2, TypographyControl } from '@spteck/react-controls-v2';
import { tokens } from '@fluentui/react-components';

type MessageBarProps = {
    children: React.ReactNode;
    type?: 'info' | 'error' | 'success' | 'warning';
    onDismiss?: () => void;
}

const MessageBar: React.FC<MessageBarProps> = ({ children, type = 'info', onDismiss }) => {
    const messageTypeStyle = () => {
        switch (type) {
            case 'error':
                return {
                    backgroundColor: tokens.colorPaletteRedBackground1,
                    borderColor: tokens.colorPaletteRedBorder2,
                    color: tokens.colorPaletteRedForeground1
                };
            case 'success':
                return {
                    backgroundColor: tokens.colorPaletteLightGreenBackground1,
                    borderColor: tokens.colorPaletteGreenBorder2,
                    color: tokens.colorPaletteGreenForeground1
                };
            case 'warning':
                return {
                    backgroundColor: tokens.colorPaletteYellowBackground1,
                    borderColor: tokens.colorPaletteYellowBorder2,
                    color: tokens.colorPaletteYellowForeground1
                };
            case 'info':
            default:
                return {
                    backgroundColor: tokens.colorNeutralBackground3,
                    borderColor: tokens.colorNeutralStroke2,
                    color: tokens.colorNeutralForeground3
                };
        }
    };

    const messageTypeIcon = () => {
        switch (type) {
            case 'error':
                return <DismissCircle24Regular style={{ color: tokens.colorPaletteRedForeground1, flexShrink: 0 }} />;
            case 'success':
                return <CheckmarkCircle24Regular style={{ color: tokens.colorPaletteGreenForeground1, flexShrink: 0 }} />;
            case 'warning':
                return <Info24Regular style={{ color: tokens.colorPaletteYellowForeground1, flexShrink: 0 }} />;
            case 'info':
            default:
                return <Info24Regular style={{ color: tokens.colorNeutralForeground3, flexShrink: 0 }} />;
        }
    };

    return (
        <StackV2 direction="horizontal" gap="s" alignItems="center" padding="m"
            style={messageTypeStyle()}>
            {messageTypeIcon()}
            <TypographyControl fontSize="xs" color={messageTypeStyle().color}>
                {children}
            </TypographyControl>
            {onDismiss &&
                <button onClick={onDismiss}
                    style={{ marginLeft: 'auto', background: 'none', border: 'none', cursor: 'pointer', padding: 0, display: 'flex', alignItems: 'center', color: 'inherit', flexShrink: 0 }}>
                    <Dismiss16Regular />
                </button>
            }
        </StackV2>
    );
}

export default MessageBar;