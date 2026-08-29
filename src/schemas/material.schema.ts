import { z } from "zod";
// const statePricesSchema = z.object({
//   "Maharashtra": z.number().min(0),
//   "Gujarat": z.number().min(0),
//   "Uttar_Pradesh": z.number().min(0),
//   "Karnataka": z.number().min(0),
//   "West_Bengal": z.number().min(0),
//   "Delhi": z.number().min(0),
//   "Odisha": z.number().min(0),
//   "Goa": z.number().min(0),
// });

const statePricesSchema = z.record(z.string(), z.number().min(0));

export const createMaterialSchema = z.object({
  name: z.string().min(1, "Material name is required"),
  code: z.string().optional(),
  materialBrandId: z.string().optional(),
  size: z.string().optional(),
  uom: z.string().optional(),
  pressure: z.string().optional(),
  hsn: z.string().optional(),
  materialTypeId: z.string().optional(),
  gst: z.string().optional(),
  purchase_price: z.number().optional(),
  sales_price: z.number().optional(),
  photos: z.array(z.string()).optional(),
  active: z.boolean().optional().default(true),
  sales_description: z.string().optional(),
  purchase_description: z.string().optional(),
  alias: z.string().optional(),
  minqty: z.number().min(0).optional(),
  date: z.string().optional(),
  vendorName: z.string().optional(),
  specification: z.string().optional(),
  category: z.string().optional(),
  stockLocation: z.string().optional(),
  quantity: z.number().min(0).optional(),
  state_prices: statePricesSchema.optional(),
});

export const updateMaterialSchema = createMaterialSchema.partial();
