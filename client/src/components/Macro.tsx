import * as React from 'react'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import SelectCustomer from './SelectCustomer'
import MacroForm from './macro-form/MacroForm'
import CustomerContext, { CustomerContextType } from '@/context/CustomerContext'
import POSummary from './po-summary/POSummary'

const macroList = [
    'kohls-towel-pdf',
    'kohls-towel',
    'kohls-rugs-pdf',
    'kohls-bedsheet',
    'walmart-bedsheet',
    'kohls-po-mismatch',
]

function Macro() {
    const { selected } = React.useContext(CustomerContext) as CustomerContextType
    return (
        <Card className="max-w-lg mx-auto">
            <CardHeader>
                <CardTitle>Macro automation</CardTitle>
            </CardHeader>
            <CardContent>
                <SelectCustomer />
                {macroList.includes(selected) && <MacroForm />}
                {selected === 'kohls-po-summary' && <POSummary />}
            </CardContent>
        </Card>
    )
}

export default Macro
