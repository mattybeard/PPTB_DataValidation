import { useCallback, useEffect, useMemo, useState } from "react";
import { createTheme, ThemeProvider, Toggle } from "@fluentui/react";
import { useConnection, useToolboxEvents } from "./hooks/useToolboxAPI";
import { TableSelection } from "./components/TableSelection";

function App() {
    const [isDarkMode, setIsDarkMode] = useState(true);
    const { connection, refreshConnection } = useConnection();

    const theme = useMemo(
        () => createTheme({
            palette: isDarkMode
                ? {
                      themePrimary: "#4aa3ff",
                      themeLighterAlt: "#03080d",
                      themeLighter: "#0c293f",
                      themeLight: "#164464",
                      themeTertiary: "#2d88c4",
                      themeSecondary: "#419cd5",
                      themeDarkAlt: "#5caeff",
                      themeDark: "#7abbff",
                      themeDarker: "#a5d2ff",
                      neutralLighterAlt: "#25292f",
                      neutralLighter: "#2c3138",
                      neutralLight: "#383f48",
                      neutralQuaternaryAlt: "#414953",
                      neutralQuaternary: "#484f59",
                      neutralTertiaryAlt: "#677281",
                      neutralTertiary: "#d9dce0",
                      neutralSecondary: "#eceef0",
                      neutralPrimaryAlt: "#f4f6f8",
                      neutralPrimary: "#ffffff",
                      neutralDark: "#fbfbfc",
                      black: "#fefefe",
                      white: "#1f2329",
                  }
                : {
                      themePrimary: "#0069c2",
                      themeLighterAlt: "#f4f9fd",
                      themeLighter: "#d0e7fa",
                      themeLight: "#a9d2f4",
                      themeTertiary: "#5ca6e6",
                      themeSecondary: "#1a81d3",
                      themeDarkAlt: "#005eae",
                      themeDark: "#004f92",
                      themeDarker: "#003a6b",
                      neutralLighterAlt: "#f8f9fb",
                      neutralLighter: "#f3f4f6",
                      neutralLight: "#e5e7ea",
                      neutralQuaternaryAlt: "#d9dde3",
                      neutralQuaternary: "#d0d4da",
                      neutralTertiaryAlt: "#b1b7c0",
                      neutralTertiary: "#8f96a3",
                      neutralSecondary: "#5f6775",
                      neutralPrimaryAlt: "#434b58",
                      neutralPrimary: "#1f2630",
                      neutralDark: "#161c24",
                      black: "#0f141b",
                      white: "#ffffff",
                  },
        }),
        [isDarkMode]
    );

    useEffect(() => {
        document.body.setAttribute("data-theme", isDarkMode ? "dark" : "light");
    }, [isDarkMode]);

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
        <ThemeProvider theme={theme}>
            <header className="header">
                <div className="headerTopRow">
                    <h1>Data Validation Tool</h1>
                    <Toggle
                        checked={isDarkMode}
                        onChange={(_, checked) => setIsDarkMode(Boolean(checked))}
                        onText="Dark"
                        offText="Light"
                        label="Theme"
                        inlineLabel
                    />
                </div>
                <p className="subtitle">Select a Dataverse table, define column tests, and run selected checks.</p>
            </header>

            <TableSelection connection={connection} />
        </ThemeProvider>
    );
}

export default App;
