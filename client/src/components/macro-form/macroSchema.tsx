import { z } from 'zod'

const macroSchema = z.object({
    mastersheet: z.instanceof(FileList),
    files: z
        .instanceof(FileList)
        .refine((files) => files?.length > 0, 'Please select at least one file.')
        .refine((files) => {
            for (const file of files) {
                const ext = file.name.split('.').pop()?.toLowerCase()
                if (ext !== 'xls' && ext !== 'xlsx' && ext !== 'xlsb' && ext !== 'pdf') {
                    return false
                }
            }
            return true
        }, 'Invalid file selected. Only xls or xlsx files are allowed.'),
})

export default macroSchema
