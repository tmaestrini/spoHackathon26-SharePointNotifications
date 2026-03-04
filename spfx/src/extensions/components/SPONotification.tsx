import { RenderDialog, StackV2, TypographyControl } from '@spteck/react-controls-v2';
import * as React from 'react'; import {
    Warning24Regular, Info24Regular,
    Checkmark24Regular, DismissCircle24Regular,
} from '@fluentui/react-icons';
import { Button, tokens } from '@fluentui/react-components';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';

export interface ISPONotificationProps {
    onClose: () => void;
    context: ListViewCommandSetContext;
}

const SPONotification: React.FC<ISPONotificationProps> = ({ onClose, context }) => {
    const [open, setOpen] = React.useState(true);

    return (
        <RenderDialog
            isOpen={open}
            dialogTitle={
                <StackV2 direction="horizontal" gap="s" alignItems="center">
                    <Warning24Regular style={{ color: tokens.colorPaletteYellowForeground2 }} />
                    <TypographyControl fontWeight="semibold" fontSize="l">
                        Confirm Action
                    </TypographyControl>
                </StackV2>
            }
            minWidth={'80vw'}

            dialogActions={
                <StackV2 paddingTop="m" direction="horizontal" gap="s"
                    justifyContent="flex-end" style={{ width: '100%' }}>
                    <Button appearance="secondary" icon={<DismissCircle24Regular />}
                        onClick={() => { setOpen(false); onClose(); }}>Cancel</Button>
                    <Button appearance="primary" icon={<Checkmark24Regular />}
                        onClick={() => { setOpen(false); onClose(); }}>Yes, Continue</Button>
                </StackV2>
            }
            onDismiss={() => { setOpen(false); onClose(); }}
        >
            <StackV2 direction="vertical" gap="l">
                <TypographyControl fontSize="s" style={{ lineHeight: '22px' }}>
                    Are you sure you want to continue? This action cannot be undone
                    and will permanently apply the changes.
                </TypographyControl>
                <StackV2 direction="horizontal" gap="s" alignItems="center" padding="m"
                    style={{
                        borderRadius: tokens.borderRadiusMedium,
                        backgroundColor: tokens.colorNeutralBackground3,
                        border: `1px solid ${tokens.colorNeutralStroke2}`
                    }}>
                    <Info24Regular style={{ color: tokens.colorNeutralForeground3, flexShrink: 0 }} />
                    <TypographyControl fontSize="xs" color={tokens.colorNeutralForeground3}>
                        This dialog is rendered using RenderDialog from @spteck/react-controls-v2.
                    </TypographyControl>
                </StackV2>
            </StackV2>
        </RenderDialog >
    );
};

export default SPONotification;
