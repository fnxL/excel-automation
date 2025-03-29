import * as React from 'react'
import { CardTitle } from './ui/card'

function FormTitle({ children }: { children: React.ReactNode }) {
    return (
        <div className="space-y-1.5 pb-6">
            <CardTitle>{children}</CardTitle>
        </div>
    )
}
export default FormTitle
