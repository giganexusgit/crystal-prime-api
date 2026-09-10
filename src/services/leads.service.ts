import {
  AlignmentType,
  BorderStyle,
  Document,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";
import { ILike } from "typeorm";
import { Clients } from "../entities/clients.entity";
import {
  FollowupStatus,
  LeadFollowup,
} from "../entities/lead-followups.entity";
import { LeadSources } from "../entities/lead-sources.entity";
import { LeadStatuses } from "../entities/lead-statuses.entity";
import { LeadTypes } from "../entities/lead-type.entity";
import { ChannelType, Leads } from "../entities/leads.entity";
import { NotificationType } from "../entities/notification.entity";
import { User } from "../entities/user.entity";
import {
  createLeadSchema,
  createMetaLeadSchema,
} from "../schemas/leads.schema";
import AppError from "../utils/appError";
import { AppDataSource } from "../utils/data-source";
import { formatQuotationDate } from "../utils/formatQuotationDate";
import { NotificationService } from "./notification.service";
import { getValidToken } from "./page-token.service";

const leadRepo = AppDataSource.getRepository(Leads);
const userRepo = AppDataSource.getRepository(User);
const leadSourceRepo = AppDataSource.getRepository(LeadSources);
const leadStatusRepo = AppDataSource.getRepository(LeadStatuses);
const leadTypeRepo = AppDataSource.getRepository(LeadTypes);
const leadFollowupRepo = AppDataSource.getRepository(LeadFollowup);
const clientRepo = AppDataSource.getRepository(Clients);

const notificationService = NotificationService();

const CELL_PADDING = {
  top: 80,
  bottom: 80,
  left: 120,
  right: 120,
};

const PARA_SPACING = {
  before: 0,
  after: 60,
};

// Create lead
export const LeadService = () => {
  // Create Lead
  const createLead = async (data: any, userData: any) => {
    const {
      first_name,
      last_name,
      company,
      phone,
      other_contact,
      email,
      location,
      budget,
      requirement,
      possibility_of_conversion,
      remark,
      source_id,
      status_id,
      type_id,
      assigned_to,
    } = data;

    // Validate email (optional now)
    if (email && typeof email !== "string") {
      throw new AppError(400, "Email must be a string");
    }

    if (phone) {
      const existingLead = await leadRepo.findOne({ where: { phone } });
      if (existingLead) {
        throw new AppError(400, "Phone Number already exists");
      }
    }

    const lead = new Leads();
    lead.first_name = first_name;
    lead.last_name = last_name;
    lead.company = company ?? "";
    lead.phone = phone ?? "";
    lead.email = email || "";
    lead.location = location ?? "";
    lead.remark = remark ?? "";
    // Handle numeric fields properly
    lead.budget = budget && budget !== "" ? Number(budget) : null;
    lead.requirement = requirement ?? "";
    lead.possibility_of_conversion =
      possibility_of_conversion && possibility_of_conversion !== ""
        ? Number(possibility_of_conversion)
        : null;
    lead.other_contact = other_contact ?? "";
    lead.created_by = `${userData?.first_name} ${userData?.last_name}`.trim();
    lead.updated_by = `${userData?.first_name} ${userData?.last_name}`.trim();

    if (source_id) {
      const source = await leadSourceRepo.findOne({ where: { id: source_id } });
      if (!source) throw new AppError(400, "Invalid Lead Source");
      lead.source = source;
    }

    if (status_id) {
      const status = await leadStatusRepo.findOne({ where: { id: status_id } });
      if (!status) throw new AppError(400, "Invalid Lead Status");
      lead.status = status;
    }

    if (type_id) {
      const type = await leadTypeRepo.findOne({ where: { id: type_id } });
      if (!type) throw new AppError(400, "Invalid Lead Type");
      lead.type = type;
    }

    // if (assigned_to) {
    //   const user = await userRepo.findOne({ where: { id: assigned_to } });
    //   if (!user) throw new AppError(400, "Invalid Assigned User");
    //   lead.assigned_to = user;
    // }

    if (assigned_to) {
      const user = await userRepo.findOne({ where: { id: assigned_to } });
      if (!user) throw new AppError(400, "Invalid Assigned User");
      lead.assigned_to = user;

      // update lastAssigned and persist
      user.lastAssigned = new Date().toISOString();
      await userRepo.save(user);
    } else {
      // auto-assign based on lead.requirement text vs user.keywords
      const requirementText = (lead.requirement || "").toString().trim();

      if (!requirementText) {
        // nothing to match against — fallback to selecting by oldest lastAssigned
        // const candidates = await userRepo.find({
        //   /* add filters: active/deleted etc. */
        // });

        const candidates = await userRepo
          .createQueryBuilder("user")
          .where("user.role_id = :roleId", {
            roleId: "a668bb29-2fbf-4a5e-be4d-9c73e990871b",
          })
          .andWhere("user.deleted = false") // if you use soft delete flag
          .getMany();

        if (!candidates || candidates.length === 0) {
          lead.assigned_to = null;
        } else {
          // treat missing lastAssigned as oldest (0) so they get priority
          candidates.sort((a, b) => {
            const ta = a.lastAssigned ? new Date(a.lastAssigned).getTime() : 0;
            const tb = b.lastAssigned ? new Date(b.lastAssigned).getTime() : 0;
            return ta - tb; // ascending => oldest first
          });

          const selected = candidates[0];
          lead.assigned_to = selected;
          selected.lastAssigned = new Date().toISOString();
          await userRepo.save(selected);
        }
      } else {
        // tokenize requirement - words and meaningful phrases
        const reqTokens = requirementText
          .toLowerCase()
          .split(/[\s,;.:\-()\/\\]+/)
          .map((t) => t.trim())
          .filter(Boolean);

        // load candidate users — restrict query as needed (active users only etc.)
        // const users = await userRepo.find(); // adapt where clause as needed

        const users = await userRepo
          .createQueryBuilder("user")
          .where("user.role_id = :roleId", {
            roleId: "a668bb29-2fbf-4a5e-be4d-9c73e990871b",
          })
          .andWhere("user.deleted = false") // if you use soft delete flag
          .getMany();

        // helper to normalize keywords stored in various formats
        const normalizeKeywords = (raw: any): string[] => {
          if (!raw) return [];
          if (Array.isArray(raw)) {
            return raw
              .map((k) => String(k).toLowerCase().trim())
              .filter(Boolean);
          }
          if (typeof raw === "string") {
            const s = raw.trim();
            if ((s.startsWith("[") && s.endsWith("]")) || s.startsWith('["')) {
              try {
                const parsed = JSON.parse(s);
                if (Array.isArray(parsed))
                  return parsed
                    .map((k) => String(k).toLowerCase().trim())
                    .filter(Boolean);
              } catch (e) {
                // fallback to CSV parse
              }
            }
            return s
              .split(/[,;|\/]+|\s+/)
              .map((k) => k.toLowerCase().trim())
              .filter(Boolean);
          }
          return String(raw)
            .toLowerCase()
            .split(/[,;|\/]+|\s+/)
            .map((k) => k.trim())
            .filter(Boolean);
        };

        // Score users by matches
        type Score = { user: any; score: number; matchedKeywords: string[] };
        const scores: Score[] = [];

        for (const u of users) {
          const kws = normalizeKeywords(u.keywords);
          if (kws.length === 0) continue;

          let score = 0;
          const matched: Set<string> = new Set();

          for (const kw of kws) {
            if (!kw) continue;
            if (reqTokens.includes(kw)) {
              score += 2;
              matched.add(kw);
              continue;
            }
            if (requirementText.toLowerCase().includes(kw)) {
              score += 1;
              matched.add(kw);
            }
          }

          if (score > 0) {
            scores.push({
              user: u,
              score,
              matchedKeywords: Array.from(matched),
            });
          }
        }

        if (scores.length === 0) {
          // NO KEYWORD MATCH -> pick user with oldest lastAssigned (round-robin-ish)
          const candidates = users; // optionally filter eligible users here
          if (!candidates || candidates.length === 0) {
            lead.assigned_to = null;
          } else {
            candidates.sort((a, b) => {
              const ta = a.lastAssigned
                ? new Date(a.lastAssigned).getTime()
                : 0;
              const tb = b.lastAssigned
                ? new Date(b.lastAssigned).getTime()
                : 0;
              return ta - tb;
            });

            const selected = candidates[0];
            lead.assigned_to = selected;
            selected.lastAssigned = new Date().toISOString();
            await userRepo.save(selected);
          }
        } else {
          // KEYWORD MATCH FOUND -> choose best, assign and update lastAssigned
          scores.sort((a, b) => {
            if (b.score !== a.score) return b.score - a.score;
            const aMatched = a.matchedKeywords.length;
            const bMatched = b.matchedKeywords.length;
            if (bMatched !== aMatched) return bMatched - aMatched;
            return 0;
          });

          const best = scores[0];
          lead.assigned_to = best.user;

          // update lastAssigned on selected user and persist
          best.user.lastAssigned = new Date().toISOString();
          await userRepo.save(best.user);

          // optional log:
          // console.log(`Auto-assigned lead to user ${best.user.id}. Matched:`, best.matchedKeywords);
        }
      }
    }

    const savedLead = await leadRepo.save(lead);

    // Send notification to assigned user if any
    if (lead.assigned_to) {
      await notificationService.createNotification(
        lead.assigned_to.id,
        NotificationType.LEAD_ASSIGNED,
        `You have been assigned a new lead: ${first_name} ${last_name}`,
        {
          leadId: savedLead.id,
          leadName: `${first_name} ${last_name}`,
          assignedBy: `${userData?.first_name} ${userData?.last_name}`,
        },
      );
    }

    return savedLead;
  };

  // Get All Leads
  const getAllLeads = async (
    filters: any = {},
    userId?: string,
    role?: string,
  ) => {
    const page = Number(filters.page) > 0 ? Number(filters.page) : 1;
    const limit = Number(filters.limit) > 0 ? Number(filters.limit) : 10;
    const skip = (page - 1) * limit;

    const {
      searchText,
      statusId,
      typeId,
      dateRange,
      referenceDate,
      followupFrom,
      followupTo,
      sourceId,
      assignedToId,
    } = filters;

    let query = leadRepo
      .createQueryBuilder("lead")
      .leftJoinAndSelect("lead.source", "source")
      .leftJoinAndSelect("lead.status", "status")
      .leftJoinAndSelect("lead.assigned_to", "assigned_to")
      .leftJoinAndSelect("lead.type", "type")
      .leftJoinAndSelect("lead.followups", "followup")
      .where("lead.deleted = false");

    // Role-based filtering - non-admins can only see their assigned leads
    // if (role && role !== "admin" && role !== "Admin") {
    //   query = query.andWhere("assigned_to.id = :userId", { userId });
    // }

    // if (role?.trim().toLowerCase() !== "admin") {
    //   if (!userId) {
    //     throw new Error("User ID is required");
    //   }
    //   const testId = "1ddc0f57-093e-4077-a4b6-3cc42df14587";
    //   query = query.andWhere("assigned_to.id = :userId", {
    //     userId: testId,
    //   });
    // }

    console.log("role", role);
    console.log("assignedToId", assignedToId);
    console.log("userId", userId);

    if (searchText && searchText.trim() !== "") {
      const search = `%${searchText.trim().toLowerCase()}%`;
      query = query.andWhere(
        `LOWER(lead.first_name) LIKE :search
         OR LOWER(lead.last_name) LIKE :search
         OR LOWER(lead.company) LIKE :search
         OR LOWER(lead.phone) LIKE :search
         OR LOWER(lead.location) LIKE :search
         OR LOWER(lead.requirement) LIKE :search
         OR LOWER(lead.email) LIKE :search`,
        { search },
      );
    }

    if (statusId && statusId !== "All Status") {
      query = query.andWhere("status.id = :statusId", { statusId });
    }

    if (typeId && typeId !== "All Type") {
      query = query.andWhere("type.id = :typeId", { typeId });
    }

    if (sourceId && sourceId !== "All Source") {
      query = query.andWhere("source.id = :sourceId", { sourceId });
    }

    if (assignedToId && assignedToId !== "All Assigned") {
      query = query.andWhere("assigned_to.id = :assignedToId", {
        assignedToId,
      });
    }

    const now = referenceDate ? new Date(referenceDate) : new Date();

    if (dateRange && dateRange !== "All") {
      let start: Date | undefined = undefined;
      let end: Date | undefined = undefined;

      if (dateRange === "Daily") {
        start = new Date(now);
        start.setHours(0, 0, 0, 0);
        end = new Date(now);
        end.setHours(23, 59, 59, 999);
      } else if (dateRange === "Weekly") {
        start = new Date(now);
        start.setDate(now.getDate() - now.getDay());
        start.setHours(0, 0, 0, 0);
        end = new Date(start);
        end.setDate(start.getDate() + 6);
        end.setHours(23, 59, 59, 999);
      } else if (dateRange === "Monthly") {
        start = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
        end = new Date(
          now.getFullYear(),
          now.getMonth() + 1,
          0,
          23,
          59,
          59,
          999,
        );
      }

      if (start && end) {
        query = query.andWhere("lead.created_at BETWEEN :start AND :end", {
          start,
          end,
        });
      }
    }

    if (followupFrom && followupTo) {
      query = query.andWhere(
        "followup.due_date BETWEEN :followupFrom AND :followupTo",
        {
          followupFrom,
          followupTo,
        },
      );
    } else if (followupFrom) {
      query = query.andWhere("followup.due_date >= :followupFrom", {
        followupFrom,
      });
    } else if (followupTo) {
      query = query.andWhere("followup.due_date <= :followupTo", {
        followupTo,
      });
    }

    query.orderBy("lead.created_at", "DESC");
    query.skip(skip).take(limit);

    const [leads, total] = await query.getManyAndCount();
    let filteredLeads = leads;

    if (role?.trim().toLowerCase() !== "admin") {
      if (!userId) {
        throw new Error("User ID is required");
      }

      filteredLeads = leads.filter(
        (lead: any) => lead.assigned_to?.id === userId,
      );
    }
    return {
      data: filteredLeads,
      pagination: {
        total,
        page,
        limit,
        totalPages: Math.ceil(total / limit),
      },
    };
  };

  // Get Lead By ID
  const getLeadById = async (id: string) => {
    const lead = await leadRepo.findOne({
      where: { id, deleted: false },
      relations: ["source", "status", "assigned_to", "type"],
    });
    if (!lead) throw new AppError(400, "Lead not found");
    return lead;
  };

  const getLeadStats = async (userId: string, role: string) => {
    // Get today's start and end timestamps in UTC
    const now = new Date();
    const today = new Date(
      Date.UTC(
        now.getUTCFullYear(),
        now.getUTCMonth(),
        now.getUTCDate(),
        0,
        0,
        0,
        0,
      ),
    );

    const isAdmin = role === "admin" || role === "Admin";
    const assignedToFilter = isAdmin ? {} : { assigned_to: { id: userId } };
    const followupUserFilter = isAdmin ? {} : { user: { id: userId } };

    const [
      totalLeads,
      assignedToMe,
      profileSent,
      convertedLeads,
      lostLeads,
      todayFollowups,
    ] = await Promise.all([
      // Total leads
      leadRepo.count({ where: { deleted: false, ...assignedToFilter } }),

      // Assigned to me
      isAdmin
        ? leadRepo.count({ where: { deleted: false } })
        : leadRepo.count({
            where: { deleted: false, assigned_to: { id: userId } },
            relations: ["assigned_to"],
          }),

      // Profile sent
      leadRepo.count({
        where: {
          deleted: false,
          status: { name: "Profile Sent" },
          ...assignedToFilter,
        },
        relations: ["status"],
      }),

      // Converted leads
      leadRepo
        .createQueryBuilder("lead")
        .leftJoin("lead.status", "status")
        .where("lead.deleted = :deleted", { deleted: false })
        .andWhere("LOWER(status.name) IN (:...statuses)", {
          statuses: ["business done", "completed", "converted to client"], // Add more if needed
        })
        .andWhere(isAdmin ? "1=1" : "lead.assigned_to = :userId", { userId })
        .getCount(),

      // Lost leads (no-interested variations)
      leadRepo
        .createQueryBuilder("lead")
        .leftJoin("lead.status", "status")
        .where("lead.deleted = :deleted", { deleted: false })
        .andWhere("LOWER(status.name) IN (:...statuses)", {
          statuses: [
            "no-interested",
            "no interested",
            "not interested",
            "no-Interested",
          ],
        })
        .andWhere(isAdmin ? "1=1" : "lead.assigned_to = :userId", { userId })
        .getCount(),

      // Today's followups (date-only comparison, exclude COMPLETED)
      leadFollowupRepo
        .createQueryBuilder("f")
        .leftJoin("f.user", "u")
        .where("f.deleted = :deleted", { deleted: false })
        .andWhere("DATE(COALESCE(f.due_date, f.created_at)) = CURRENT_DATE")
        .andWhere("(f.status IS NULL OR f.status != :completed)", {
          completed: FollowupStatus.COMPLETED,
        })
        .andWhere(isAdmin ? "1=1" : "u.id = :userId", { userId })
        .getCount(),
    ]);

    return {
      totalLeads,
      assignedToMe,
      profileSent,
      convertedLeads,
      lostLeads,
      todayFollowups,
    };
  };

  // Update Lead
  const updateLead = async (id: string, data: any, userData: any) => {
    const {
      first_name,
      last_name,
      company,
      phone,
      other_contact,
      email,
      location,
      budget,
      requirement,
      possibility_of_conversion,
      remark,
      source_id,
      status_id,
      type_id,
      assigned_to,
    } = data;

    const lead = await leadRepo.findOne({
      where: { id, deleted: false },
      relations: ["assigned_to"],
    });
    if (!lead) throw new AppError(400, "Lead not found");

    if (email) {
      lead.email = email || "";
    }

    lead.first_name = first_name ?? lead.first_name;
    lead.last_name = last_name ?? lead.last_name;
    lead.company = company ?? lead.company;
    lead.phone = phone ?? lead.phone;
    lead.location = location ?? lead.location;
    lead.budget = budget ?? lead.budget;
    lead.requirement = requirement ?? lead.requirement;
    lead.remark = remark ?? lead.remark;
    lead.possibility_of_conversion =
      possibility_of_conversion ?? lead.possibility_of_conversion;
    lead.other_contact = other_contact ?? lead.other_contact;
    lead.updated_by = `${userData?.first_name} ${userData?.last_name}`.trim();

    if (source_id !== undefined) {
      lead.source =
        source_id === null
          ? null
          : await leadSourceRepo.findOne({ where: { id: source_id } });
    }

    if (status_id) {
      const status = await leadStatusRepo.findOne({ where: { id: status_id } });
      if (status) {
        lead.status = status;

        //Status is completed add lead into client table.
        const currentStatus = status?.name?.toLocaleLowerCase();
        if (
          currentStatus === "business done" ||
          currentStatus === "converted to client"
        ) {
          const existingLead = await clientRepo.findOne({
            where: {
              lead: { id: lead.id },
            },
          });
          //if not already exist then create
          if (!existingLead) {
            const name = (lead.first_name ?? "") + (lead.last_name ?? "");

            let email = lead?.email || "";

            const contact_number = lead?.phone ?? "";
            const address = lead?.location ?? "";
            const company_name = lead?.company ?? "";
            const leadId = lead.id;

            const client = clientRepo.create({
              name,
              email,
              lead: { id: leadId },
              contact_number,
              address,
              company_name,
              contact_person: name,
            });
            await clientRepo.save(client); //save lead to client.
          }
        }
      }
    }

    if (type_id !== undefined) {
      lead.type =
        type_id === null
          ? null
          : await leadTypeRepo.findOne({ where: { id: type_id } });
    }

    // Handle lead escalation
    if (data.escalate_to === true) {
      lead.escalate_to = true;

      // Create a duplicate lead
      const duplicateLead = leadRepo.create({
        first_name: lead.first_name,
        last_name: lead.last_name,
        company: lead.company,
        phone: lead.phone,
        other_contact: lead.other_contact,
        email: lead.email,
        location: lead.location,
        budget: lead.budget,
        requirement: lead.requirement,
        possibility_of_conversion: lead.possibility_of_conversion,
        remark: lead.remark,
        channel: lead.channel,
        source: lead.source,
        type: lead.type,
        status: lead.status,
        assigned_to: null, // Unassigned initially
        created_by: `${userData?.first_name} ${userData?.last_name}`.trim(),
        updated_by: `${userData?.first_name} ${userData?.last_name}`.trim(),
        escalate_to: false, // Reset escalation flag for duplicate
      });

      await leadRepo.save(duplicateLead);

      // Notify all admins about the escalated lead
      const adminUsers = await userRepo.find({
        where: { role: { role: "admin" }, deleted: false },
        relations: ["role"],
      });

      for (const admin of adminUsers) {
        await notificationService.createNotification(
          admin.id,
          NotificationType.LEAD_ESCALATED,
          `Lead Escalated: ${lead.first_name} ${lead.last_name} (${
            lead.phone || lead.email
          }) - Duplicate created`,
          {
            leadId: lead.id,
            duplicateLeadId: duplicateLead.id,
            leadName: `${lead.first_name} ${lead.last_name}`,
            leadContact: lead.phone || lead.email,
            escalatedBy: `${userData?.first_name} ${userData?.last_name}`,
            requirement: lead.requirement,
          },
        );
      }
    }

    if (assigned_to !== undefined && assigned_to !== lead.assigned_to?.id) {
      lead.assigned_to =
        assigned_to === null
          ? null
          : await userRepo.findOne({ where: { id: assigned_to } });

      await notificationService.createNotification(
        assigned_to,
        NotificationType.LEAD_ASSIGNED,
        `You have been assigned a new lead: ${lead.first_name} ${lead.last_name}`,
        {
          leadId: lead.id,
          leadName: `${lead.first_name} ${lead.last_name}`,
          assignedBy: `${userData?.first_name} ${userData?.last_name}`,
        },
      );
    }

    const savedLead = await leadRepo.save(lead);

    // // Send notification to assigned user if any
    // if (lead.assigned_to) {
    //   await notificationService.createNotification(
    //     lead.assigned_to.id,
    //     NotificationType.LEAD_ASSIGNED,
    //     `You have been assigned a new lead: ${first_name} ${last_name}`,
    //     {
    //       leadId: savedLead.id,
    //       leadName: `${first_name} ${last_name}`,
    //       assignedBy: `${userData?.first_name} ${userData?.last_name}`,
    //     }
    //   );
    // }

    return savedLead;
  };

  // Soft Delete Lead
  const softDeleteLead = async (id: string) => {
    const lead = await leadRepo.findOne({
      where: { id },
      relations: ["source", "status", "assigned_to", "type"],
    });

    if (!lead) throw new AppError(400, "Lead not found");

    lead.deleted = true;
    lead.deleted_at = new Date();

    await leadRepo.save(lead);

    return {
      status: "success",
      message: "Lead soft deleted successfully",
      data: lead,
    };
  };

  // get of no leads assigned to user
  const getTodayAssignedLeadsCount = async (userId: string) => {
    const now = new Date();

    const startOfToday = new Date(
      now.getFullYear(),
      now.getMonth(),
      now.getDate(),
    );

    const endOfToday = new Date(
      now.getFullYear(),
      now.getMonth(),
      now.getDate() + 1,
    );

    return await leadRepo
      .createQueryBuilder("lead")
      .where("lead.deleted = false")
      .andWhere("lead.assigned_to = :userId", { userId })
      .andWhere("lead.created_at >= :startOfToday", { startOfToday })
      .andWhere("lead.created_at < :endOfToday", { endOfToday })
      .getCount();
  };

  //  Export Leads to Excel
  const exportLeadsToExcel = async (
    userId: string,
    userRole: string,
    searchText?: string,
    statusId?: string,
    typeId?: string,
    dateRange?: "All" | "Daily" | "Weekly" | "Monthly",
    referenceDate?: Date,
    followupFrom?: Date,
    followupTo?: Date,
    sourceId?: string,
    assignedToId?: string,
  ): Promise<ExcelJS.Workbook> => {
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
      page: 1,
      limit: 10000, // Large limit to get all leads for export
    };

    const result = await getAllLeads(filters, userId, userRole);
    const leads = result.data;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Leads");

    worksheet.columns = [
      { header: "Sr No", key: "sr_no", width: 6 },
      { header: "First Name", key: "first_name", width: 20 },
      { header: "Last Name", key: "last_name", width: 20 },
      { header: "Company", key: "company", width: 25 },
      { header: "Phone", key: "phone", width: 20 },
      { header: "Other Contact", key: "other_contact", width: 20 },
      { header: "Email", key: "email", width: 30 },
      { header: "Location", key: "location", width: 20 },
      { header: "Budget", key: "budget", width: 15 },
      { header: "Requirement", key: "requirement", width: 40 },
      {
        header: "Possibility of Conversion (%)",
        key: "possibility_of_conversion",
        width: 25,
      },
      {
        header: "Remark",
        key: "remark",
        width: 25,
      },
      { header: "Source", key: "source", width: 20 },
      { header: "Status", key: "status", width: 20 },
      { header: "type", key: "type", width: 20 },
      { header: "Assigned To", key: "assigned_to", width: 25 },
      { header: "Created At", key: "created_at", width: 25 },
    ];

    leads.forEach((lead, index) => {
      worksheet.addRow({
        sr_no: index + 1,
        first_name: lead.first_name,
        last_name: lead.last_name,
        company: lead.company ?? "",
        phone: lead.phone ?? "",
        other_contact: lead.other_contact ?? "",
        email: lead.email ?? "",
        location: lead.location ?? "",
        budget: lead.budget ?? 0,
        requirement: lead.requirement ?? "",
        remark: lead.remark ?? "",
        possibility_of_conversion: lead.possibility_of_conversion ?? "",
        source: lead.source?.name ?? "",
        status: lead.status?.name ?? "",
        type: lead.type?.name ?? "",
        assigned_to: `${lead.assigned_to?.first_name} ${lead.assigned_to?.last_name}`,
        created_at: lead.created_at?.toLocaleString() ?? "",
      });
    });

    return workbook;
  };

  const generateLeadTemplate = async (): Promise<ExcelJS.Workbook> => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Leads");

    worksheet.columns = [
      { header: "first_name", key: "first_name", width: 20 },
      { header: "last_name", key: "last_name", width: 20 },
      { header: "company", key: "company", width: 25 },
      { header: "phone", key: "phone", width: 15 },
      { header: "other_contact", key: "other_contact", width: 15 },
      { header: "email", key: "email", width: 25 },
      { header: "location", key: "location", width: 20 },
      { header: "requirement", key: "requirement", width: 20 },
      { header: "budget", key: "budget", width: 15 },
      {
        header: "possibility_of_conversion",
        key: "possibility_of_conversion",
        width: 25,
      },
      { header: "remark", key: "remark", width: 15 },
      { header: "source", key: "source", width: 15 },
      { header: "status", key: "status", width: 15 },
      { header: "type", key: "type", width: 15 },
    ];

    return workbook;
  };

  // Service to handle Excel upload
  const uploadLeadsFromExcelService = async (
    fileBuffer: Buffer,
    user: User,
  ) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(fileBuffer as any);
    const worksheet = workbook.worksheets[0];

    const headers: string[] = [];
    worksheet.getRow(1).eachCell((cell) => {
      headers.push(cell.text.toLowerCase().trim());
    });

    // Define required fields
    const requiredFields = ["first_name", "last_name"];
    const missingFields = requiredFields.filter(
      (field) => !headers.includes(field),
    );
    if (missingFields.length > 0) {
      throw new AppError(
        400,
        `Missing required fields: ${missingFields.join(", ")}`,
      );
    }

    const leadsToInsert: any[] = [];
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row

      const leadData: any = {};
      headers.forEach((header, colIndex) => {
        const cell = row.getCell(colIndex + 1);
        const cellValue = cell.value;

        let value = "";
        if (
          cellValue &&
          typeof cellValue === "object" &&
          "text" in cellValue &&
          cellValue.text &&
          typeof cellValue.text === "object" &&
          "richText" in (cellValue.text as any) &&
          Array.isArray((cellValue.text as any).richText)
        ) {
          value = ((cellValue.text as any).richText as any[])
            .map((rt: any) => rt.text)
            .join("");
        } else if (
          cellValue &&
          typeof cellValue === "object" &&
          "text" in cellValue &&
          typeof (cellValue as any).text === "string"
        ) {
          value = (cellValue as any).text;
        } else {
          value = cell.text || "";
        }
        leadData[header] = value;
      });

      leadData._rowNumber = rowNumber; // Attach row number for error tracking
      leadsToInsert.push(leadData);
    });

    const savedLeads = [];

    for (const data of leadsToInsert) {
      const rowNumber = data._rowNumber;

      // Check if email already exists (only if email is provided and valid)
      const email = data.email || "";
      const emailString = String(email).trim();

      if (emailString && emailString.length > 0) {
        // Basic email validation
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(emailString)) {
          throw new AppError(
            400,
            `Invalid email format at row ${rowNumber}: ${emailString}`,
          );
        }

        const existingEmail = await leadRepo
          .createQueryBuilder("lead")
          .where("lead.deleted = false")
          .andWhere("lead.email = :email", { email: emailString })
          .getOne();

        if (existingEmail) {
          throw new AppError(
            400,
            `Email already exists at row ${rowNumber}: ${emailString}`,
          );
        }
      }

      // Create lead object
      const lead = leadRepo.create({
        first_name: data.first_name || "",
        last_name: data.last_name || "",
        company: data.company || "",
        phone: data.phone || "",
        other_contact: data.other_contact || "",
        email: emailString,
        location: data.location || "",
        budget: Number(data.budget) || 0,
        requirement: data.requirement || "",
        remark: data.remark || "",
        possibility_of_conversion:
          Number(data.possibility_of_conversion) || null,
        created_by: `${user.first_name} ${user.last_name}` || "",
        updated_by: `${user.first_name} ${user.last_name}` || "",
      });

      // Find Source by Name
      if (data.source) {
        const source = await leadSourceRepo.findOne({
          where: { name: String(data.source).trim() },
        });
        if (!source) {
          throw new AppError(
            400,
            `Invalid source name at row ${rowNumber}: ${data.source}`,
          );
        }
        lead.source = source;
      }

      if (data.type) {
        const type = await leadTypeRepo.findOne({
          where: { name: String(data.type).trim() },
        });
        if (!type) {
          throw new AppError(
            400,
            `Invalid type name at row ${rowNumber}: ${data.type}`,
          );
        }
        lead.type = type;
      }

      // Find Status by Name
      if (data.status) {
        const status = await leadStatusRepo.findOne({
          where: { name: String(data.status).trim() },
        });
        if (!status) {
          throw new AppError(
            400,
            `Invalid status name at row ${rowNumber}: ${data.status}`,
          );
        }
        lead.status = status;
      }

      lead.assigned_to = user;

      const saved = await leadRepo.save(lead);
      savedLeads.push(saved);
    }

    return { total: savedLeads.length, leads: savedLeads };
  };

  const findLeadByEmail = async ({ email }: { email: string }) => {
    return await leadRepo
      .createQueryBuilder("lead")
      .where("lead.deleted = false")
      .andWhere("lead.email = :email", { email })
      .getOne();
  };

  const findLeadByPhoneNumber = async ({ phone }: { phone: string }) => {
    return await leadRepo.findOne({
      where: { phone, deleted: false },
    });
  };

  // Group leads by status for a given date range and user
  const groupLeadsByStatus = async (
    dateRange: "Daily" | "Weekly" | "Monthly" | "Yearly",
    userId?: string,
    role?: string,
    referenceDate?: Date,
  ) => {
    const now = referenceDate ? new Date(referenceDate) : new Date();
    let start: Date | undefined;
    let end: Date | undefined;
    if (dateRange === "Weekly") {
      start = new Date(now);
      start.setDate(now.getDate() - now.getDay());
      start.setHours(0, 0, 0, 0);
      end = new Date(start);
      end.setDate(start.getDate() + 6);
      end.setHours(23, 59, 59, 999);
    } else if (dateRange === "Daily") {
      start = new Date(now);
      start.setHours(0, 0, 0, 0);
      end = new Date(now);
      end.setHours(23, 59, 59, 999);
    } else if (dateRange === "Monthly") {
      start = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
      end = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59, 999);
    } else if (dateRange === "Yearly") {
      start = new Date(now.getFullYear(), 0, 1, 0, 0, 0, 0);
      end = new Date(now.getFullYear(), 11, 31, 23, 59, 59, 999);
    }
    const qb = leadRepo
      .createQueryBuilder("lead")
      .leftJoin("lead.status", "status")
      .select(["status.name AS status", "COUNT(*)::int AS count"])
      .where("lead.deleted = false");
    if (start && end) {
      qb.andWhere("lead.created_at BETWEEN :start AND :end", { start, end });
    }
    if (role !== "admin" && role !== "Admin" && userId) {
      qb.andWhere("lead.assigned_to = :userId", { userId });
    }
    return await qb.groupBy("status.name").getRawMany();
  };

  // daily leads for a given date range and user

  const getDailyLeadStats = async (userId: string, role?: string) => {
    const start = new Date();
    start.setHours(0, 0, 0, 0);

    const end = new Date();
    end.setHours(23, 59, 59, 999);

    const convertedStatuses = [
      "business done",
      "completed",
      "converted to client",
    ];

    // Base query for assigned leads
    const leadQb = leadRepo
      .createQueryBuilder("lead")
      .leftJoin("lead.status", "status")
      .where("lead.deleted = false");

    if (role !== "admin" && role !== "Admin") {
      leadQb.andWhere("lead.assigned_to = :userId", { userId });
    }

    const [todayLeads, convertedLeads, followupsDone, pendingFollowups] =
      await Promise.all([
        // Today's leads
        leadQb
          .clone()
          .andWhere("lead.created_at BETWEEN :start AND :end", { start, end })
          .getCount(),

        // Converted leads today
        leadQb
          .clone()
          .andWhere("lead.created_at BETWEEN :start AND :end", { start, end })
          .andWhere("LOWER(status.name) IN (:...statuses)", {
            statuses: convertedStatuses,
          })
          .getCount(),

        // Today's completed followups
        AppDataSource.getRepository(LeadFollowup)
          .createQueryBuilder("followup")
          .leftJoin("followup.lead", "lead")
          .where("followup.deleted = false")
          .andWhere("followup.status = :status", {
            status: "COMPLETED",
          })
          .andWhere("followup.created_at BETWEEN :start AND :end", {
            start,
            end,
          })
          .andWhere(
            role !== "admin" && role !== "Admin"
              ? "lead.assigned_to = :userId"
              : "1=1",
            { userId },
          )
          .getCount(),

        // Pending followups
        AppDataSource.getRepository(LeadFollowup)
          .createQueryBuilder("followup")
          .leftJoin("followup.lead", "lead")
          .where("followup.deleted = false")
          .andWhere("followup.status = :status", {
            status: "PENDING",
          })
          .andWhere(
            role !== "admin" && role !== "Admin"
              ? "lead.assigned_to = :userId"
              : "1=1",
            { userId },
          )
          .getCount(),
      ]);

    return {
      todayLeads,
      followupsDone,
      pendingFollowups,
      convertedLeads,
    };
  };
  // Group leads by type for a given date range and user
  const groupLeadsByType = async (
    dateRange: "Weekly" | "Monthly" | "Yearly",
    userId?: string,
    role?: string,
    referenceDate?: Date,
  ) => {
    const now = referenceDate ? new Date(referenceDate) : new Date();
    let start: Date | undefined;
    let end: Date | undefined;
    if (dateRange === "Weekly") {
      start = new Date(now);
      start.setDate(now.getDate() - now.getDay());
      start.setHours(0, 0, 0, 0);
      end = new Date(start);
      end.setDate(start.getDate() + 6);
      end.setHours(23, 59, 59, 999);
    } else if (dateRange === "Monthly") {
      start = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
      end = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59, 999);
    } else if (dateRange === "Yearly") {
      start = new Date(now.getFullYear(), 0, 1, 0, 0, 0, 0);
      end = new Date(now.getFullYear(), 11, 31, 23, 59, 59, 999);
    }
    const qb = leadRepo
      .createQueryBuilder("lead")
      .leftJoin("lead.type", "type")
      .select(["type.name AS type", "COUNT(*)::int AS count"])
      .where("lead.deleted = false");
    if (start && end) {
      qb.andWhere("lead.created_at BETWEEN :start AND :end", { start, end });
    }
    if (role !== "admin" && role !== "Admin" && userId) {
      qb.andWhere("lead.assigned_to = :userId", { userId });
    }
    return await qb.groupBy("type.name").getRawMany();
  };

  const verifyWebhook = (mode: any, token: any): boolean => {
    const VERIFY_TOKEN = process.env.META_VERIFY_TOKEN;

    if (mode === "subscribe" && token === VERIFY_TOKEN) {
      return true;
    }
    return false;
  };

  const handleMetaLead = async (leadId: string, channel: ChannelType) => {
    const PAGE_ACCESS_TOKEN = await getValidToken();
    const META_DATA_SOURCE_ENDPOINT = process.env.META_DATA_SOURCE_ENDPOINT!;

    // Step 1: Fetch the lead details
    const url = `${META_DATA_SOURCE_ENDPOINT}/${leadId}?access_token=${PAGE_ACCESS_TOKEN}`;
    const response = await fetch(url);

    if (!response.ok) {
      throw new AppError(400, "Failed to fetch lead from Meta");
    }

    const data = await response.json();
    const fieldData = data.field_data;

    const mapped: Record<string, any> = {};

    for (const item of fieldData) {
      const value = item.values?.[0];
      switch (item.name) {
        case "email":
          mapped.email = value;
          break;
        case "attachments":
          mapped.attachments = item.values;
          break;
        default:
          mapped[item.name] = value;
          break;
      }
    }

    createMetaLeadSchema.parse(mapped);

    // Step 2: Fetch Campaign Name
    let campaignName: string | null = null;
    try {
      const adId = data.ad_id;
      if (adId) {
        const adUrl = `${META_DATA_SOURCE_ENDPOINT}/${adId}?fields=adset{campaign{name}}&access_token=${PAGE_ACCESS_TOKEN}`;
        const adResp = await fetch(adUrl);

        if (adResp.ok) {
          const adData = await adResp.json();
          campaignName = adData?.adset?.campaign?.name || null;
        }
      }
    } catch (err) {
      console.error("\n\nFailed to fetch campaign info: ", err, "\n\n");
    }

    // Step 3: Find Type in DB (optional)
    let leadType: LeadTypes | null = null;

    try {
      if (campaignName) {
        leadType = await leadTypeRepo.findOne({
          where: { name: ILike(campaignName) },
        });

        if (!leadType) {
          console.log("\n\n\nNo matching Campaign name\n\n\n");
          return;
        }
      }
    } catch (e) {
      console.log("\n\n\nError while fetching lead types", e, "\n\n\n");
    }

    let phoneNumber = null;

    if (mapped?.phone_number) {
      phoneNumber = mapped.phone_number;
    } else if (mapped?.phone) {
      phoneNumber = mapped.phone;
    }

    // Step 4: Save lead
    const newLead = leadRepo.create({
      first_name: mapped.first_name,
      last_name: mapped.last_name,
      company: mapped.company,
      phone: phoneNumber,
      other_contact: mapped.other_contact ?? null,
      email: mapped.email || "",
      location: mapped.address,
      budget:
        mapped.budget && mapped.budget !== ""
          ? parseFloat(mapped.budget)
          : null,
      requirement: mapped.requirement,
      attachments: mapped.attachments || [],
      channel,
      type: leadType || null,
    });

    await leadRepo.save(newLead);
  };

  const handleGoogleLead = async (payload: any, receivedApiKey: string) => {
    const expectedApiKey = process.env.GOOGLE_SECRETE_KEY;

    if (!expectedApiKey || receivedApiKey !== expectedApiKey) {
      throw new AppError(401, "Unauthorized: Invalid API Key");
    }

    if (!payload) {
      throw new AppError(400, "Invalid payload from Google");
    }

    const prepData = {
      ...payload,
      email: payload.email || "",
      budget:
        payload.budget && payload.budget !== ""
          ? parseInt(payload.budget)
          : null,

      attachments: Array.isArray(payload.attachments)
        ? payload.attachments
        : payload.attachments
          ? [payload.attachments]
          : [],
    };

    const data = createLeadSchema.parse(prepData);

    const newLead = leadRepo.create({
      first_name: data.first_name,
      last_name: data.last_name,
      company: data.company,
      phone: data.phone,
      other_contact: data.other_contact,
      email: data.email || "",
      location: data.location,
      budget: data.budget && data.budget !== "" ? Number(data.budget) : null,
      requirement: data.requirement,
      attachments: data.attachments,
      channel: ChannelType.GOOGLE,
    });

    await leadRepo.save(newLead);
  };

  const generateQuotationDocService = async (
    leadId: string,
    products: any[] = [],
    productsText: string,
    subtotal: number,
    taxPercent: number,
    finalAmount: number,
    proposalDate?: string,
    proposalNumber?: string,
    proposalText?: string,
  ) => {
    const lead = await leadRepo.findOne({
      where: { id: leadId },
      relations: ["assigned_to", "status", "type", "source"],
    });

    if (!lead) throw new AppError(404, "Lead not found");

    const logoCandidates = [
      path.join(__dirname, "../../public/logo.jpeg"),
      path.join(__dirname, "../../../src/public/logo.jpeg"),
      path.join(process.cwd(), "src/public/logo.jpeg"),
      path.join(process.cwd(), "public/logo.jpeg"),
    ];

    const resolvedLogoPath = logoCandidates.find((p) => fs.existsSync(p));
    if (!resolvedLogoPath) {
      throw new AppError(500, "Quotation logo not found");
    }

    const logoBuffer = fs.readFileSync(resolvedLogoPath);

    const doc = new Document({
      creator: "Crystal Prime",
      title: "Quotation Document",
      sections: [
        {
          children: [
            /* LOGO */
            new Paragraph({
              spacing: { after: 300 },
              children: [
                new ImageRun({
                  data: logoBuffer,
                  type: "png",
                  transformation: { width: 180, height: 100 },
                }),
              ],
            }),

            /* DIVIDER */
            new Paragraph({
              border: {
                bottom: { style: BorderStyle.SINGLE, size: 4, color: "C0C0C0" },
              },
              spacing: { after: 400 },
            }),

            /* FROM / TO TABLE */
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              borders: FULL_TABLE_BORDER(),
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      width: { size: 50, type: WidthType.PERCENTAGE }, // ✅ 50%

                      margins: CELL_PADDING,
                      children: [
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({
                              text: "Quotation From:",
                              bold: true,
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({
                              text: "Crystal Portal Cabins",
                              bold: true,
                              size: 22,
                            }),
                          ],
                        }),

                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({
                              text: "Email ID: sales@crytalprime.com",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({
                              text: "Mobile No: +91 9022236505",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({
                              text: "GST No: 27AAJFC0916G1ZO",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({
                              text: "Address: Shop No. 1, House 136/02, Adivali Bhutavali, Near L & T Infotech, Mahape Shil Road, Mahape, Navi Mumbai, Maharashtra 400710, India",
                            }),
                          ],
                        }),
                      ],
                    }),

                    new TableCell({
                      width: { size: 50, type: WidthType.PERCENTAGE }, // ✅ 50%

                      margins: CELL_PADDING,
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.RIGHT,
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({ text: "Quotation To:", bold: true }),
                          ],
                        }),
                        new Paragraph({
                          alignment: AlignmentType.RIGHT,
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({
                              text: lead.company || "-",
                              bold: true,
                              size: 22,
                            }),
                          ],
                        }),
                        new Paragraph({
                          alignment: AlignmentType.RIGHT,
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({
                              text: `Client Name: ${lead.first_name || ""} ${lead.last_name || ""}`,
                            }),
                          ],
                        }),

                        new Paragraph({
                          alignment: AlignmentType.RIGHT,
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({
                              text: `Client Email ID: ${lead.email || "-"}`,
                            }),
                          ],
                        }),

                        new Paragraph({
                          alignment: AlignmentType.RIGHT,
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({
                              text: `Client Mobile No: ${lead.phone || "-"}`,
                            }),
                          ],
                        }),

                        new Paragraph({
                          alignment: AlignmentType.RIGHT,
                          spacing: PARA_SPACING,
                          children: [new TextRun({ text: "Client GST No: -" })],
                        }),

                        new Paragraph({
                          alignment: AlignmentType.RIGHT,
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({
                              text: `Client Address: ${lead.location || "-"}`,
                            }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),

            new Paragraph({ spacing: { after: 120 } }),

            /* DATE & NUMBER */
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      margins: CELL_PADDING,
                      borders: BORDER_BOX(),
                      children: [
                        new Paragraph({
                          spacing: PARA_SPACING,
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: "Proposal Date: ",
                              bold: true,
                            }),
                            new TextRun({
                              text:
                                formatQuotationDate(proposalDate) || "_____",
                            }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      margins: CELL_PADDING,
                      borders: BORDER_BOX(),
                      children: [
                        new Paragraph({
                          spacing: PARA_SPACING,
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: "Proposal Number: ",
                              bold: true,
                            }),
                            new TextRun({ text: proposalNumber || "_____" }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),

            new Paragraph({ spacing: { after: 120 } }),

            /* PROPOSAL DETAILS */
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      margins: CELL_PADDING,
                      borders: BORDER_BOX(),
                      children: [
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({
                              text: "Proposal Details:",
                              bold: true,
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({ text: proposalText || "-" }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),

            new Paragraph({ spacing: { after: 120 } }),

            /* PRODUCTS TABLE */
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              borders: FULL_TABLE_BORDER(),
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      margins: CELL_PADDING,
                      children: [
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({ text: "Product Name", bold: true }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      margins: CELL_PADDING,
                      children: [
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({
                              text: "Price Per Quantity (₹)",
                              bold: true,
                            }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      margins: CELL_PADDING,
                      children: [
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [new TextRun({ text: "Size", bold: true })],
                        }),
                      ],
                    }),
                    new TableCell({
                      margins: CELL_PADDING,
                      children: [
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({ text: "Quantity", bold: true }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      margins: CELL_PADDING,
                      children: [
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({ text: "State", bold: true }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      margins: CELL_PADDING,
                      children: [
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({ text: "Total", bold: true }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),

                ...products.map(
                  (item) =>
                    new TableRow({
                      children: [
                        new TableCell({
                          margins: CELL_PADDING,
                          children: [
                            new Paragraph({
                              spacing: PARA_SPACING,
                              children: [
                                new TextRun({ text: item.name || "-" }),
                              ],
                            }),
                          ],
                        }),
                        new TableCell({
                          margins: CELL_PADDING,
                          children: [
                            new Paragraph({
                              spacing: PARA_SPACING,
                              children: [
                                new TextRun({
                                  text: String(item.salePrice || 0),
                                }),
                              ],
                            }),
                          ],
                        }),
                        new TableCell({
                          margins: CELL_PADDING,
                          children: [
                            new Paragraph({
                              spacing: PARA_SPACING,
                              children: [
                                new TextRun({
                                  text: String(item.productSize || 0),
                                }),
                              ],
                            }),
                          ],
                        }),
                        new TableCell({
                          margins: CELL_PADDING,
                          children: [
                            new Paragraph({
                              spacing: PARA_SPACING,
                              children: [
                                new TextRun({ text: String(item.count || 0) }),
                              ],
                            }),
                          ],
                        }),
                        new TableCell({
                          margins: CELL_PADDING,
                          children: [
                            new Paragraph({
                              spacing: PARA_SPACING,
                              children: [
                                new TextRun({
                                  text: String(item.state || "-"),
                                }),
                              ],
                            }),
                          ],
                        }),
                        new TableCell({
                          margins: CELL_PADDING,
                          children: [
                            new Paragraph({
                              spacing: PARA_SPACING,
                              children: [
                                new TextRun({
                                  text: String(item.totalPrice || 0),
                                }),
                              ],
                            }),
                          ],
                        }),
                      ],
                    }),
                ),
              ],
            }),

            new Paragraph({ spacing: { after: 120 } }),

            /* PRODUCT DESCRIPTION */
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      margins: CELL_PADDING,
                      borders: BORDER_BOX(),
                      children: [
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({
                              text: "Products Description:",
                              bold: true,
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({ text: productsText || "-" }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),

            new Paragraph({ spacing: { after: 120 } }),

            /* PRICE SUMMARY */
            new Table({
              width: { size: 60, type: WidthType.PERCENTAGE },
              alignment: AlignmentType.RIGHT,
              borders: FULL_TABLE_BORDER(),
              rows: [
                ["Subtotal", subtotal],
                [`Tax (${taxPercent}%)`, (subtotal * taxPercent) / 100],
                ["Final Amount", finalAmount],
              ].map(
                ([label, value], i) =>
                  new TableRow({
                    children: [
                      new TableCell({
                        margins: CELL_PADDING,
                        children: [
                          new Paragraph({
                            spacing: PARA_SPACING,
                            children: [
                              new TextRun({
                                text: String(label),
                                bold: i === 2,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        margins: CELL_PADDING,
                        children: [
                          new Paragraph({
                            spacing: PARA_SPACING,
                            children: [
                              new TextRun({
                                text: `₹ ${Number(value).toFixed(2)}`,
                                bold: i === 2,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
              ),
            }),

            // Terms And Condtions
            // After your existing totals table, add a spacer paragraph and then the terms table:

            new Paragraph({ spacing: PARA_SPACING }),

            new Paragraph({
              spacing: PARA_SPACING,
              children: [
                new TextRun({
                  text: "OTHER TERMS & CONDITIONS: -",
                  bold: true,
                  underline: {},
                }),
              ],
              alignment: AlignmentType.CENTER,
            }),

            new Table({
              width: { size: 9000, type: WidthType.DXA },
              alignment: AlignmentType.CENTER,
              borders: FULL_TABLE_BORDER(),
              rows: [
                {
                  label: "Freight & Challan",
                  value: "Freight cost & RTO fine will be extra at actual.",
                },
                {
                  label: "Taxes",
                  value: "GST @ 18% extra as per mention above.",
                },
                {
                  label: "Delivery Schedule",
                  value:
                    "7 to 8 working days from the date of receipt of Purchase Order & Advance.",
                },
                {
                  label: "Loading",
                  value: "Will be in our scope.",
                },
                {
                  label: "Unloading",
                  value: "To be borne by the client.",
                },
                {
                  label: "Warranty",
                  value: "365 days against any manufacturing defects.",
                },
                {
                  label: "Payment Terms",
                  value:
                    "75% Advance along with purchase order and balance after inspection of the material before dispatch.",
                },
                {
                  label: "OFFER VALIDITY",
                  value: "Prices mentioned are valid for a period of 1 month.",
                },
              ].map(
                ({ label, value }) =>
                  new TableRow({
                    children: [
                      new TableCell({
                        margins: CELL_PADDING,
                        width: { size: 2500, type: WidthType.DXA },
                        children: [
                          new Paragraph({
                            spacing: PARA_SPACING,
                            numbering: { reference: "bullets", level: 0 },
                            children: [
                              new TextRun({
                                text: label,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        margins: CELL_PADDING,
                        width: { size: 6500, type: WidthType.DXA },
                        children: [
                          new Paragraph({
                            spacing: PARA_SPACING,
                            children: [
                              new TextRun({
                                text: value,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
              ),
            }),

            // bank details
            new Paragraph({ spacing: PARA_SPACING }),

            new Paragraph({
              spacing: PARA_SPACING,
              children: [
                new TextRun({
                  text: "Details of Company",
                  bold: true,
                  underline: {},
                }),
              ],
              alignment: AlignmentType.CENTER,
            }),

            new Table({
              width: { size: 9000, type: WidthType.DXA },
              alignment: AlignmentType.CENTER,
              borders: FULL_TABLE_BORDER(),
              rows: [
                {
                  label: "Account Name",
                  value: "Crystal Cabin",
                },
                {
                  label: "Bank",
                  value: "HDFC Bank",
                },
                {
                  label: "Account Type",
                  value: "Current Account",
                },
                {
                  label: "Account No.",
                  value: "50200053443660",
                },
                {
                  label: "Branch",
                  value: "Navi Mumbai",
                },
                {
                  label: "IFSC",
                  value: "HDFC0001602",
                },
                {
                  label: "GST No.",
                  value: "27AAJFC0916G1ZO",
                },
                {
                  label: "Pan No.",
                  value: "AAJFC0916G",
                },
                {
                  label: "MSME No.",
                  value: "MH33A0205607",
                },
              ].map(
                ({ label, value }) =>
                  new TableRow({
                    children: [
                      new TableCell({
                        margins: CELL_PADDING,
                        width: { size: 4500, type: WidthType.DXA },
                        children: [
                          new Paragraph({
                            spacing: PARA_SPACING,
                            children: [
                              new TextRun({
                                text: label,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        margins: CELL_PADDING,
                        width: { size: 4500, type: WidthType.DXA },
                        children: [
                          new Paragraph({
                            spacing: PARA_SPACING,
                            children: [
                              new TextRun({
                                text: value,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
              ),
            }),

            // Other Specification
            new Paragraph({ spacing: PARA_SPACING }),

            new Paragraph({
              spacing: PARA_SPACING,
              children: [
                new TextRun({
                  text: "TECHNICAL SPECIFICATION FOR NEW PREFAB CABIN",
                  bold: true,
                  underline: {},
                }),
              ],
              alignment: AlignmentType.CENTER,
            }),

            new Table({
              width: { size: 9000, type: WidthType.DXA },
              alignment: AlignmentType.CENTER,
              borders: FULL_TABLE_BORDER(),
              rows: [
                // Header Row
                new TableRow({
                  tableHeader: true,
                  children: [
                    new TableCell({
                      margins: CELL_PADDING,
                      width: { size: 800, type: WidthType.DXA },
                      // shading: { fill: "D3D3D3", type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({ text: "Sr.No.", bold: true }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      margins: CELL_PADDING,
                      width: { size: 2000, type: WidthType.DXA },
                      // shading: { fill: "D3D3D3", type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({ text: "Components", bold: true }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      margins: CELL_PADDING,
                      width: { size: 6200, type: WidthType.DXA },
                      // shading: { fill: "D3D3D3", type: ShadingType.CLEAR },
                      children: [
                        new Paragraph({
                          spacing: PARA_SPACING,
                          children: [
                            new TextRun({ text: "Descriptions", bold: true }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
                // Data Rows
                ...[
                  {
                    no: "1",
                    component: "Bottom Frame",
                    description:
                      "100X50 formed MS I-Beam ISI Marked {Not Applicable}",
                  },
                  {
                    no: "2",
                    component: "Stiffener Bottom",
                    description:
                      "80X40 & 50X50 mm Square Pipes. ISI Marked {Not Applicable}",
                  },
                  {
                    no: "3",
                    component: "Top Frame",
                    description:
                      "Frame made of 50X50 Square Pipes. ISI Marked {Not Applicable}",
                  },
                  {
                    no: "4",
                    component: "Stiffener Top",
                    description:
                      "50X50 & 40X20 Square pipes. ISI Marked {Not Applicable}",
                  },
                  {
                    no: "5",
                    component: "Side Stiffener",
                    description:
                      "Corner Post Shall be of 60X60 Square Pipes & wall thickness of 60mm. ISI Marked {Not Applicable}",
                  },
                  {
                    no: "6",
                    component: "Paneling Outside",
                    description:
                      "Specially Corrugated G.I./M.S. 1.2 mm Thick Sheet Welded on MS frame. The steel sheet shall be treated for corrosion resistance. Panels shall be vertically corrugated, sheet shall be continuously welded to top side and base structure so as to offer better strength to weight ratio. All gaps will be sealed at edges and at seams, ISI Marked (UTTAM/JSW/ESSAR) {Not Applicable}",
                  },
                  {
                    no: "7",
                    component: "Internal Wall Paneling",
                    description:
                      "8mm Thick MDF Make of Green Panel & all Vertically & Horizontal Corners Will be Neatly & Smoothly Finished with Aluminum Sections & L Angles.",
                  },
                  {
                    no: "8",
                    component: "Toilet Wall Paneling",
                    description:
                      "Toilet Walls Shall be of 4mm Hardener Sheets Water Roof & All Vertically & Horizontal Corners Shall be Neatly & Smoothly Finished with Aluminum Sections.",
                  },
                  {
                    no: "9",
                    component: "Roof Outside",
                    description:
                      "1.2 mm thick GI/M.S. Sheet Properly Sloped & Water Tight Protect from Anti Rust with Gutter System ISI Marked {Not Applicable}",
                  },
                  {
                    no: "10",
                    component: "False Ceiling",
                    description:
                      "9mm thick MDF Make of Green Panel & All Vertically & Horizontal Corners Will be Neatly & Smoothly Finished With Aluminum Sections & L Angles.",
                  },
                  {
                    no: "11",
                    component: "Bottom Flooring",
                    description:
                      "On the bottom frame 18 mm Thick V-Board I.e. Cemented Bonded Fiber Sheets Shall (Visaka or Bison Panel) be Fixed by Means of Self-Taping Screw {Not Applicable} & 1mm Thick PVC Vinyl Carpet Shall be Fixed on The Panel. (100% water, Termite & Dimensional Proof with Vinyl Carpet)",
                  },
                  {
                    no: "12",
                    component: "Aluminum Sliding Windows",
                    description:
                      "Double Shutter Sliding Aluminum Powder Coated Windows of Thickness 1.2mm with 4mm Thick Glass for All Windows, Safety Grills from Outside & Canopy On The Top Of Windows. The Door Shall be Of External Opening Type Made Out Of The Same Material as Wall Panels & Canopy Above the Door.",
                  },
                  {
                    no: "13",
                    component: "Main Door",
                    description:
                      "Frame Work by Tubular Pipe Of 30X30 & The Door Interior Shall be Finished With Same Material Matching With That Of The Cabin Interior With Insulations. Door With Stand Hardware Locks & Handles Al drop Lock",
                  },
                  {
                    no: "14",
                    component: "Insulation",
                    description:
                      "50 mm Thick Glass Wool Density Of 64 kg/m3 With Insulation For Top & 25mm Side Walls To Avoid Heat",
                  },
                  {
                    no: "15",
                    component: "Wiring",
                    description:
                      "All Wiring Shall be Concealed & Shall be PVC Insulated Copper Wires Of ISI Quality, Suitable for 240 Volts, 50 HZ Single Phase AC Power Supply. With MCB Protection Split Air Conditioner Point Along With Separate MCB. ISI Marked Main Supply, MCB & AC :- 4mm, Fan & Light :- 1.5 mm, Neutral :- 2.5mm, Sockets :- 2.5mm, Earthling :- 1mm",
                  },
                  {
                    no: "16",
                    component: "Outside Painting",
                    description:
                      "All Components Are Painted With 2 Coats Of BERGER/ESDEE Epoxy Primer & 2 Coats Of Corrosion Free BERGER/ESDEE Paint i.e. Synthetic Epoxy Paint.",
                  },
                  {
                    no: "17",
                    component: "Hooks for Cabin Lifting",
                    description:
                      "Specially Formed Hooks For Easy Lift & Shift. {Not Applicable}",
                  },
                  {
                    no: "18",
                    component: "Ear thing",
                    description: "Ear thing point shall be provided for safety",
                  },
                  {
                    no: "19",
                    component: "Legs",
                    description:
                      "6 inch height, six legs at all four corners & center {Not Applicable}",
                  },
                  {
                    no: "20",
                    component: "Furniture's",
                    description:
                      "All Furniture's Shall be made of 18mm Pre-Laminated (Particle) Board.",
                  },
                ].map(
                  ({ no, component, description }) =>
                    new TableRow({
                      children: [
                        new TableCell({
                          margins: CELL_PADDING,
                          width: { size: 800, type: WidthType.DXA },
                          children: [
                            new Paragraph({
                              spacing: PARA_SPACING,
                              children: [new TextRun({ text: no })],
                            }),
                          ],
                        }),
                        new TableCell({
                          margins: CELL_PADDING,
                          width: { size: 2000, type: WidthType.DXA },
                          children: [
                            new Paragraph({
                              spacing: PARA_SPACING,
                              children: [
                                new TextRun({ text: component, bold: true }),
                              ],
                            }),
                          ],
                        }),
                        new TableCell({
                          margins: CELL_PADDING,
                          width: { size: 6200, type: WidthType.DXA },
                          children: [
                            new Paragraph({
                              spacing: PARA_SPACING,
                              children: [new TextRun({ text: description })],
                            }),
                          ],
                        }),
                      ],
                    }),
                ),
              ],
            }),
          ],
        },
      ],
    });

    return await Packer.toBuffer(doc);
  };

  // Helper for borders
  function FULL_TABLE_BORDER() {
    return {
      top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      insideHorizontal: {
        style: BorderStyle.SINGLE,
        size: 1,
        color: "000000",
      },
      insideVertical: {
        style: BorderStyle.SINGLE,
        size: 1,
        color: "000000",
      },
    };
  }

  function BORDER_BOX() {
    return {
      top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
    };
  }

  // Helper
  function ALL_BORDERS() {
    return {
      top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
    };
  }

  return {
    generateQuotationDocService,
    handleGoogleLead,
    handleMetaLead,
    verifyWebhook,
    createLead,
    getAllLeads,
    getLeadStats,
    getDailyLeadStats,
    getLeadById,
    updateLead,
    softDeleteLead,
    exportLeadsToExcel,
    generateLeadTemplate,
    uploadLeadsFromExcelService,
    findLeadByEmail,
    findLeadByPhoneNumber,
    groupLeadsByStatus,
    groupLeadsByType,
    getTodayAssignedLeadsCount,
  };
};
