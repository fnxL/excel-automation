import * as React from 'react'
import { Check, ChevronsUpDown } from 'lucide-react'
import { cn } from '@/lib/utils'
import { Button } from '@/components/ui/button'
import {
    Command,
    CommandEmpty,
    CommandGroup,
    CommandInput,
    CommandItem,
    CommandList,
} from '@/components/ui/command'
import { Popover, PopoverContent, PopoverTrigger } from '@/components/ui/popover'
import CustomerContext, { CustomerContextType } from '@/context/CustomerContext'

function SelectCustomer() {
    const [open, setOpen] = React.useState(false)
    const { customers, selected, updateSelected } = React.useContext(
        CustomerContext
    ) as CustomerContextType

    return (
        <Popover open={open} onOpenChange={setOpen}>
            <PopoverTrigger asChild>
                <Button
                    variant="outline"
                    role="combobox"
                    aria-expanded={open}
                    className="w-full justify-between"
                >
                    {selected
                        ? customers.find((customer) => customer.value === selected)?.label
                        : 'Select customer...'}
                    <ChevronsUpDown className="opacity-50" />
                </Button>
            </PopoverTrigger>
            <PopoverContent className="w-[28.9rem] pt-0">
                <Command>
                    <CommandInput placeholder="Search customer..." />
                    <CommandList>
                        <CommandEmpty>No customer found.</CommandEmpty>
                        <CommandGroup>
                            {customers.map((customer) => (
                                <CommandItem
                                    key={customer.value}
                                    value={customer.value}
                                    onSelect={(currentValue) => {
                                        updateSelected(
                                            currentValue === selected ? '' : currentValue
                                        )
                                        setOpen(false)
                                    }}
                                >
                                    {customer.label}
                                    <Check
                                        className={cn(
                                            'ml-auto',
                                            selected === customer.value
                                                ? 'opacity-100'
                                                : 'opacity-0'
                                        )}
                                    />
                                </CommandItem>
                            ))}
                        </CommandGroup>
                    </CommandList>
                </Command>
            </PopoverContent>
        </Popover>
    )
}

export default SelectCustomer
