import * as React from 'react'
import CustomerContext, { Customer } from './CustomerContext'

const CustomerProvider = ({ children }: { children: React.ReactNode }) => {
    const [selected, setSelected] = React.useState('')
    const [customers, setCustomers] = React.useState<Customer[]>([])

    React.useEffect(() => {
        async function fetchCustomers() {
            const response = await fetch(`${import.meta.env.VITE_BASE_API_URL}/customers`)

            if (response.ok) {
                const data = await response.json()
                setCustomers(data.customers)
                console.log(data)
            } else {
                throw new Error('Error fetching customers')
            }
        }

        fetchCustomers()
    }, [])

    function updateSelected(customer: string) {
        setSelected(customer)
    }

    return (
        <CustomerContext.Provider value={{ customers, selected, updateSelected }}>
            {children}
        </CustomerContext.Provider>
    )
}

export default CustomerProvider
