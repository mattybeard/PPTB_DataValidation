import { useCallback } from "react";
import { ConnectionStatus } from "./components/ConnectionStatus";
import { useConnection, useToolboxEvents } from "./hooks/useToolboxAPI";
import { TableSelection } from "./components/TableSelection";

function App() {
    const { connection, isLoading, refreshConnection } = useConnection();

    // Handle platform events
    const handleEvent = useCallback(
        (event: string, _data: any) => {
            switch (event) {
                case 'connection:updated':
                case 'connection:created':
                case 'connection:deleted':
                    refreshConnection();
                    break;

                case 'terminal:output':
                case 'terminal:command:completed':
                case 'terminal:error':
                    // Terminal events handled by dedicated components
                    break;
            }
        },
        [refreshConnection]
    );

    useToolboxEvents(handleEvent);

    return (
        <>
            <header className="header">
                <h1>Data Validation Tool</h1>
                <p className="subtitle">Select a Dataverse table, define column tests, and run selected checks.</p>
            </header>

            <ConnectionStatus connection={connection} isLoading={isLoading} />

            <TableSelection connection={connection} />
        </>
    );
}

export default App;
