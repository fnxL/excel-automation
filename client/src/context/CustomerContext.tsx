import * as React from 'react'

export type CustomerContextType = {
    selected: string
    updateSelected: (customer: string) => void
    customers: Customer[]
}

export interface Customer {
    label: string
    value: string
}

const CustomerContext = React.createContext<CustomerContextType | null>(null)

export default CustomerContext
