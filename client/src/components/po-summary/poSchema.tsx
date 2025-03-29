import { z } from 'zod'

const poSchema = z.object({
    mastersheet: z.instanceof(FileList),
    poNumbers: z.string(),
})

export default poSchema
