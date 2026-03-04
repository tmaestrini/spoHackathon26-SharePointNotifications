import * as React from 'react';

export interface ISPONotificationProps {
    onClose: () => void;
}

const SPONotification: React.FC<ISPONotificationProps> = ({ onClose }) => {
    return (
        <div>
            <h2>Notification Dialog</h2>
            <p>This is a notification.</p>
            <button onClick={onClose}>Close</button>
        </div>
    );
};

export default SPONotification;
