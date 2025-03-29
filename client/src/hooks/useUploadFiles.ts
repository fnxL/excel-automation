import { useState } from 'react'

function getFileNameFromHeader(header: string) {
    const match = header.match(/filename\*?=([^;]+)/i)
    if (match) {
        return match[1].replace(/['"]/g, '').trim()
    }
    return ''
}

function useUploadFiles() {
    const [isLoading, setIsLoading] = useState(false)
    const [error, setError] = useState('')

    async function uploadFiles(formData: FormData, url: string) {
        const customer = formData.get('customerName')?.toString()
        console.log(`Customer: ${customer}`)
        try {
            setIsLoading(true)
            const response = await fetch(url, {
                method: 'POST',
                body: formData,
            })
            if (!response.ok) {
                const error = await response.json()
                setError(`Error: ${error.detail}`)
                return
            }
            const contentHeader = response.headers.get('content-disposition')
            let fileName: string = 'default.zip'
            if (customer) fileName = customer
            if (contentHeader) fileName = getFileNameFromHeader(contentHeader)
            console.log(Object.fromEntries(response.headers))

            // Create a hidden <a> element to trigger download
            const link = document.createElement('a')
            link.href = window.URL.createObjectURL(await response.blob())
            link.setAttribute('download', fileName)
            document.body.appendChild(link)
            link.click()
            link.remove()
            setIsLoading(false)
        } catch (error) {
            setError(`Error: ${error}`)
            setIsLoading(false)
        }
    }
    return {
        isLoading,
        error,
        uploadFiles,
    }
}

export default useUploadFiles
