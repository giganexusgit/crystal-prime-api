import { NextFunction, Request, Response } from "express";
import { ChannelType } from "../entities/leads.entity";
import {
  createIndiaMartLeadSchema,
  createLeadSchema,
  updateLeadSchema,
} from "../schemas/leads.schema";
import { LeadService } from "../services/leads.service";
import { findUserById } from "../services/user.service";

const service = LeadService();

export const leadController = () => {
  // Create Lead
  const createLead = async (
    req: Request,
    res: Response,
    next: NextFunction,
  ) => {
    try {
      const userData = res?.locals?.user;
      const parsed = createLeadSchema.parse(req.body);

      // Handle email conversion - can be string or array
      if (parsed.email) {
        // Ensure email is a string, trim whitespace
        parsed.email = String(parsed.email).trim();
      }

      const result = await service.createLead(parsed, userData);

      res
        .status(201)
        .json({ status: "success", message: "Lead created", data: result });
    } catch (error) {
      next(error);
    }
  };

  // Get All Leads
  const getAllLeads = async (
    req: Request,
    res: Response,
    next: NextFunction,
  ) => {
    try {
      const userId = res?.locals?.user?.id;
      const role = res?.locals?.user?.role?.role;
      const searchText = req.query.searchText as string | undefined;
      const statusId = req.query.statusId as string | undefined;
      const typeId = req.query.typeId as string | undefined;
      const dateRange = req.query.dateRange as
        | ("All" | "Daily" | "Weekly" | "Monthly")
        | undefined;
      const referenceDate = req.query.referenceDate
        ? new Date(req.query.referenceDate as string)
        : undefined;
      const followupFrom = req.query.followupFrom
        ? new Date(req.query.followupFrom as string)
        : undefined;
      const followupTo = req.query.followupTo
        ? new Date(req.query.followupTo as string)
        : undefined;
      const sourceId = req.query.sourceId as string | undefined;
      const assignedToId = req.query.assignedToId as string | undefined;
      const page = req.query.page ? parseInt(req.query.page as string) : 1;
      const limit = req.query.limit ? parseInt(req.query.limit as string) : 10;

      const filters = {
        searchText,
        statusId,
        typeId,
        dateRange,
        referenceDate,
        followupFrom,
        followupTo,
        sourceId,
        assignedToId,
        page,
        limit,
      };
      console.log("My ROle", role);

      const result = await service.getAllLeads(filters, userId, role);
      const leadStats = await service.getLeadStats(userId, role);

      res.status(200).json({
        status: "success",
        message: "All Leads fetched",
        data: {
          list: result.data,
          pagination: result.pagination,
          stats: leadStats,
        },
      });
    } catch (error) {
      next(error);
    }
  };

  // Get Lead by ID
  const getLeadById = async (
    req: Request,
    res: Response,
    next: NextFunction,
  ) => {
    try {
      const { id } = req.params;
      const result = await service.getLeadById(id);
      res
        .status(200)
        .json({ status: "success", message: "Lead fetched", data: result });
    } catch (error) {
      next(error);
    }
  };

  // Update Lead
  const updateLead = async (
    req: Request,
    res: Response,
    next: NextFunction,
  ) => {
    try {
      const userData = res?.locals?.user;
      const { id } = req.params;
      const parsed = updateLeadSchema.parse(req.body);

      if (parsed.email) {
        // Ensure email is a string, trim whitespace
        parsed.email = String(parsed.email).trim();
      }

      const result = await service.updateLead(id, parsed, userData);

      res
        .status(200)
        .json({ status: "success", message: "Lead updated", data: result });
    } catch (error) {
      next(error);
    }
  };

  // Soft Delete Lead
  const softDeleteLead = async (
    req: Request,
    res: Response,
    next: NextFunction,
  ) => {
    try {
      const { id } = req.params;
      const result = await service.softDeleteLead(id);
      res
        .status(200)
        .json({ status: "success", message: "Lead deleted", data: result });
    } catch (error) {
      next(error);
    }
  };

  // Export Leads to Excel
  const exportLeadsExcelController = async (req: Request, res: Response) => {
    try {
      const userId = res.locals.user.id;
      const userData = await findUserById(userId);
      const userRole = userData.role.role;

      // Collect all filter params
      const searchText = req.query.searchText as string | undefined;
      const statusId = req.query.statusId as string | undefined;
      const typeId = req.query.typeId as string | undefined;
      const dateRange = req.query.dateRange as
        | ("All" | "Daily" | "Weekly" | "Monthly")
        | undefined;
      const referenceDate = req.query.referenceDate
        ? new Date(req.query.referenceDate as string)
        : undefined;
      const followupFrom = req.query.followupFrom
        ? new Date(req.query.followupFrom as string)
        : undefined;
      const followupTo = req.query.followupTo
        ? new Date(req.query.followupTo as string)
        : undefined;
      const sourceId = req.query.sourceId as string | undefined;
      const assignedToId = req.query.assignedToId as string | undefined;

      const workbook = await service.exportLeadsToExcel(
        userId,
        userRole,
        searchText,
        statusId,
        typeId,
        dateRange,
        referenceDate,
        followupFrom,
        followupTo,
        sourceId,
        assignedToId,
      );

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      );
      res.setHeader(
        "Content-Disposition",
        `attachment; filename=leads_${Date.now()}.xlsx`,
      );

      await workbook.xlsx.write(res);
      res.end();
    } catch (error) {
      console.error("Error exporting leads:", error);
      res.status(500).json({ message: "Failed to export leads" });
    }
  };

  // Download Lead Template
  const downloadLeadTemplate = async (req: Request, res: Response) => {
    try {
      const workbook = await service.generateLeadTemplate();

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      );
      res.setHeader(
        "Content-Disposition",
        "attachment; filename=leads_template.xlsx",
      );

      await workbook.xlsx.write(res);
      res.end();
    } catch (error) {
      console.error("Error generating template:", error);
      res.status(500).json({ message: "Failed to download the template" });
    }
  };

  // Upload Leads from Excel
  const uploadLeadsFromExcel = async (
    req: Request,
    res: Response,
    next: NextFunction,
  ) => {
    try {
      const user = res.locals.user;
      if (!req.file) {
        return res
          .status(400)
          .json({ status: "error", message: "No file uploaded" });
      }

      const result = await service.uploadLeadsFromExcelService(
        req.file.buffer,
        user,
      );
      res.status(201).json({
        status: "success",
        message: "Leads uploaded successfully",
        data: result,
      });
    } catch (error) {
      next(error);
    }
  };

  const verifyMetaWebhook = (req: Request, res: Response) => {
    const mode = req.query["hub.mode"];
    const token = req.query["hub.verify_token"];
    const challenge = req.query["hub.challenge"];

    const verified = service.verifyWebhook(mode, token);

    if (verified) {
      res.status(200).send(challenge);
    } else {
      res.status(403).send("Forbidden");
    }
  };

  const metaLeadWebhook = async (
    req: Request,
    res: Response,
    next: NextFunction,
  ) => {
    try {
      const body = req.body;
      if (
        (body.object === "page" || body.object === "instagram") &&
        body.entry &&
        body.entry[0]?.changes &&
        body.entry[0].changes[0]?.field === "leadgen"
      ) {
        const leadgenData = body.entry[0].changes[0].value;
        const leadId = leadgenData.leadgen_id;

        let channel = ChannelType.FACEBOOK;

        if (body.object === "instagram") {
          channel = ChannelType.INSTAGRAM;
        }
        await service.handleMetaLead(leadId, channel);
        res.status(200).json({ status: "success", message: "Lead processed" });
      } else {
        res
          .status(400)
          .json({ status: "error", message: "Invalid webhook payload" });
      }
    } catch (error) {
      console.log("Lead webhook error: ", error);
      next(error);
    }
  };

  const googleLeadWebhook = async (
    req: Request,
    res: Response,
    next: NextFunction,
  ) => {
    try {
      const receivedApiKey = req.headers["x-api-key"] as string;
      const payload = req.body;

      await service.handleGoogleLead(payload, receivedApiKey);

      res.status(200).json({
        status: "success",
        message: "Google lead processed successfully",
      });
    } catch (err) {
      next(err);
    }
  };

  const exportQuotationDoc = async (
    req: Request,
    res: Response,
    next: NextFunction,
  ) => {
    try {
      const { id } = req.params;

      // Validate request body
      // const validatedData = generateQuotationSchema.parse(req.body);
      const validatedData = req.body;
      const {
        proposalDate,
        proposalNumber,
        proposalText,
        products,
        productsText,
        subtotal,
        taxPercent,
        finalAmount,
      } = validatedData;

      const buffer = await service.generateQuotationDocService(
        id,

        products,
        productsText,
        subtotal,
        taxPercent,
        finalAmount,
        new Date(proposalDate).toISOString(),
        proposalNumber,
        proposalText,
      );

      // Build file name: quotation_<lead-name>.docx
      const lead = await service.getLeadById(id);
      const leadNameRaw = `${lead.first_name || ""} ${
        lead.last_name || ""
      }`.trim();
      const safeLeadName = (leadNameRaw || id)
        .replace(/\s+/g, "_")
        .replace(/[^a-zA-Z0-9_-]/g, "");

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      );
      res.setHeader(
        "Content-Disposition",
        `attachment; filename=quotation_${safeLeadName}.docx`,
      );
      res.send(buffer);
    } catch (error) {
      next(error);
    }
  };

  const INDIA_MART_API_KEY = process.env.INDIAMART_API_KEY;

  const indiaMartLeadWebhook = async (req: Request, res: Response) => {
    try {
      const payload = req.body;
      const userData = res?.locals?.user;

      console.log("payload", payload);
      // 🔁 FIELD MAPPING (IndiaMART → Your Schema)
      const leadData = {
        first_name: payload.RESPONSE.SENDER_NAME?.split(" ")[0] || "Unknown",
        last_name: payload.RESPONSE.SENDER_NAME?.split(" ").slice(1).join(" "),
        phone: payload.RESPONSE.SENDER_MOBILE,
        email: payload.RESPONSE.SENDER_EMAIL,
        company: payload.RESPONSE.SENDER_COMPANY,
        location: `${payload.RESPONSE.SENDER_CITY}, ${payload.RESPONSE.SENDER_STATE}`,
        requirement:
          payload.RESPONSE.QUERY_MESSAGE || payload.RESPONSE.QUERY_PRODUCT_NAME,
        source_id: "d8f50868-045a-40d2-ab57-36ccb91c129b",
        possibility_of_conversion: 50,
      };

      const validatedLead = createIndiaMartLeadSchema.parse(leadData);

      await service.createLead(validatedLead, userData);

      return res.status(200).json({
        message: "Lead received successfully",
        lead_id: "lead.id",
      });
    } catch (error: any) {
      console.error("IndiaMART Webhook Error:", error);
      return res.status(400).json({
        message: "Invalid lead data",
        error: error.message,
      });
    }
  };

  return {
    exportQuotationDoc,
    googleLeadWebhook,
    metaLeadWebhook,
    verifyMetaWebhook,
    createLead,
    getAllLeads,
    getLeadById,
    updateLead,
    softDeleteLead,
    exportLeadsExcelController,
    downloadLeadTemplate,
    uploadLeadsFromExcel,
    indiaMartLeadWebhook,
  };
};
