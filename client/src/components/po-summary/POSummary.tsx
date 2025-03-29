import * as React from 'react'
import CustomerContext from '@/context/CustomerContext'
import { CustomerContextType } from '@/context/CustomerContext'
import { useForm } from 'react-hook-form'
import { zodResolver } from '@hookform/resolvers/zod'
import { z } from 'zod'
import { Separator } from '@radix-ui/react-separator'
import FormTitle from '../FormTitle'
import { Form, FormField, FormItem, FormLabel, FormMessage, FormControl } from '../ui/form'
import { Input } from '../ui/input'
import { Textarea } from '../ui/textarea'
import useUploadFiles from '@/hooks/useUploadFiles'
import { Button } from '../ui/button'
import { Loader2 } from 'lucide-react'
import poSchema from './poSchema'

function POSummary() {
    const { customers, selected } = React.useContext(CustomerContext) as CustomerContextType
    const customerLabel = customers.find((customer) => customer.value === selected)?.label
    const { uploadFiles, isLoading, error } = useUploadFiles()

    const form = useForm<z.infer<typeof poSchema>>({
        resolver: zodResolver(poSchema),
    })

    function onSubmit(values: z.infer<typeof poSchema>) {
        const formData = new FormData()
        formData.append('mastersheet', values.mastersheet[0])
        formData.append('customerName', selected)
        formData.append('poNumbers', values.poNumbers)
        const poSummaryURL = `${import.meta.env.VITE_BASE_API_URL}/posummary?customer=${selected}`
        uploadFiles(formData, poSummaryURL)
        form.reset()
    }

    return (
        <>
            <Separator className="my-6" />
            <FormTitle>{customerLabel}</FormTitle>
            <Form {...form}>
                <form onSubmit={form.handleSubmit(onSubmit)}>
                    <FormField
                        name="mastersheet"
                        control={form.control}
                        render={({ field }) => (
                            <FormItem>
                                <FormLabel>Mastersheet:</FormLabel>
                                <FormControl>
                                    <Input
                                        accept="application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                        type="file"
                                        onChange={(e) => field.onChange(e.target.files)}
                                    />
                                </FormControl>
                                <FormMessage />
                            </FormItem>
                        )}
                    />
                    <FormField
                        control={form.control}
                        name="poNumbers"
                        render={({ field }) => (
                            <FormItem>
                                <FormLabel>PO Numbes</FormLabel>
                                <FormControl>
                                    <div className="grid w-full gap-4">
                                        <Textarea placeholder="Enter PO Numbers" {...field} />
                                        {isLoading ? (
                                            <Button>
                                                <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                                                Please wait{' '}
                                            </Button>
                                        ) : (
                                            <Button type="submit" disabled={isLoading}>
                                                Submit
                                            </Button>
                                        )}
                                    </div>
                                </FormControl>
                                <FormMessage />
                            </FormItem>
                        )}
                    />

                    {error && <p className="text-red-500">{error}</p>}
                </form>
            </Form>
        </>
    )
}

export default POSummary
