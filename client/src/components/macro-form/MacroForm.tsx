import CustomerContext, { CustomerContextType } from '@/context/CustomerContext'
import * as React from 'react'
import { Form, FormField, FormItem, FormLabel, FormMessage, FormControl } from '../ui/form'
import { useForm } from 'react-hook-form'
import macroSchema from './macroSchema'
import { zodResolver } from '@hookform/resolvers/zod'
import { z } from 'zod'
import FormTitle from '../FormTitle'
import { Input } from '../ui/input'
import { Separator } from '@/components/ui/separator'
import useUploadFiles from '@/hooks/useUploadFiles'
import { Button } from '../ui/button'
import { Loader2 } from 'lucide-react'

function MacroForm() {
    const { customers, selected } = React.useContext(CustomerContext) as CustomerContextType
    const { uploadFiles, isLoading, error } = useUploadFiles()
    const customerLabel = customers.find((customer) => customer.value === selected)?.label
    const form = useForm<z.infer<typeof macroSchema>>({
        resolver: zodResolver(macroSchema),
    })

    function onSubmit(values: z.infer<typeof macroSchema>) {
        const formData = new FormData()
        Array.from(values.files).forEach((file) => {
            formData.append('files', file)
        })
        formData.append('mastersheet', values.mastersheet[0])
        formData.append('customerName', selected)
        console.log(values)
        const macroURL = `${import.meta.env.VITE_BASE_API_URL}/generate/macro?customer=${selected}`
        uploadFiles(formData, macroURL)
        form.reset()
    }

    return (
        <>
            <Separator className="my-6" />
            <FormTitle>{customerLabel}</FormTitle>
            <Form {...form}>
                <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-8">
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
                        name="files"
                        control={form.control}
                        render={({ field }) => (
                            <FormItem>
                                <FormLabel>PO Files:</FormLabel>
                                <FormControl>
                                    <Input
                                        accept="application/pdf, application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel.sheet.binary.macroenabled.12"
                                        type="file"
                                        multiple
                                        onChange={(e) => field.onChange(e.target.files)}
                                    />
                                </FormControl>
                                <FormMessage />
                            </FormItem>
                        )}
                    />
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
                    {error && <p className="text-red-500">{error}</p>}
                </form>
            </Form>
        </>
    )
}

export default MacroForm
