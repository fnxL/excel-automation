import { ThemeProvider } from '@/components/theme-provider'
import Header from './components/Header'
import Macro from './components/Macro'
import CustomerProvider from './context/CustomerProvider'

function App() {
    return (
        <ThemeProvider defaultTheme="dark" storageKey="vite-ui-theme">
            <CustomerProvider>
                <Header>Automation</Header>
                <Macro />
            </CustomerProvider>
        </ThemeProvider>
    )
}

export default App
