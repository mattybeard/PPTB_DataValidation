import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import './index.css';
import { ServiceProvider, ServiceProviderContext } from './framework/ServiceProvider';
import { ViewModel } from './models/ViewModel';

// Load in framework componenets for ease of re-use later on
const serviceProvider = new ServiceProvider();
const vm = new ViewModel();
serviceProvider.register("vm", vm);

// Ensure DOM is ready and root element exists
const rootElement = document.getElementById('root');
if (rootElement && !rootElement.hasAttribute('data-reactroot-initialized')) {
    // Mark as initialized to prevent double rendering
    rootElement.setAttribute('data-reactroot-initialized', 'true');
    
    createRoot(rootElement).render(
        <ServiceProviderContext.Provider value={serviceProvider}>
            <StrictMode>
                <App />
            </StrictMode>
        </ServiceProviderContext.Provider>
    );
} else if (!rootElement) {
    console.error('Root element not found. Make sure the HTML contains <div id="root"></div>');
}
