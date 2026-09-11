import { Request, Response, NextFunction } from "express";
import { LeadService } from "../services/leads.service";
import { ProjectService } from "../services/projects.service";
import { ProjectTaskService } from "../services/project-task.service";
import { getEILogChartData } from "../services/eilog.service";
import { ClientFollowupService } from "../services/clients-followups.service";
import { findAllUsers } from "../services/user.service";
import { Leads } from "../entities/leads.entity";
import { AppDataSource } from "../utils/data-source";

const leadService = LeadService();
const projectService = ProjectService();
const projectTaskService = ProjectTaskService();
const clientFollowupService = ClientFollowupService();
export const dashboardController = () => {
  const getDashboardSummary = async (
    req: Request,
    res: Response,
    next: NextFunction,
  ): Promise<void> => {
    const leadRepository = AppDataSource.getRepository(Leads);
    try {
      const user = res?.locals?.user;
      const userId: string = user?.id;
      const role: string = user?.role?.role;

      // Admin: full dashboard
      if (role === "admin") {
        // Fetch all required data in parallel
        const [
          leadStats,
          projectStatusCounts,

          leadStatusWeekly,
          leadStatusMonthly,
          leadStatusYearly,
          leadTypeWeekly,
          leadTypeMonthly,
          leadTypeYearly,
          allProjects,
          yearlyChart,
          monthlyChart,
          weeklyChart,
          // todayClientFollowupsCount,
          todayLeadsFollowupsCount,
        ] = await Promise.all([
          leadService.getLeadStats(userId, role),
          projectService.getProjectStatusCounts(userId, role),
          leadService.groupLeadsByStatus("Weekly", userId, role),
          leadService.groupLeadsByStatus("Monthly", userId, role),
          leadService.groupLeadsByStatus("Yearly", userId, role),
          leadService.groupLeadsByType("Weekly", userId, role),
          leadService.groupLeadsByType("Monthly", userId, role),
          leadService.groupLeadsByType("Yearly", userId, role),
          projectService.getAllProjectDashboard(userId, role),
          getEILogChartData(userId, role, "yearly"),
          getEILogChartData(userId, role, "monthly"),
          getEILogChartData(userId, role, "weekly"),
          // clientFollowupService.getTodayFollowupsCount(userId, role),
          leadService.getTodayAssignedLeadsCount(userId),
        ]);

        // Stats for cards
        const stats = [
          {
            count: String(leadStats.totalLeads || 0),
            title: "Total Leads",
            subtitle: "Over All leads",
          },
          {
            count: String(todayLeadsFollowupsCount || 0),
            title: "Assigned Leads",
            subtitle: "Today's Assigned Leads",
          },
          {
            count: String(leadStats.convertedLeads || 0),
            title: "Converted Leads",
            subtitle: "Weekly Leads",
          },
          // {
          //   count: String(leadStats.lostLeads || 0),
          //   title: "Lost Leads",
          //   subtitle: "Weekly Leads"
          // },
          {
            count:
              leadStats.totalLeads > 0
                ? `${Math.round(
                    (leadStats.convertedLeads / leadStats.totalLeads) * 100,
                  )}%`
                : "0%",
            title: "Conversion Rate",
            subtitle: "Lead to Customer",
          },
        ];
        console.log("projectStatusCounts", projectStatusCounts);

        // Staff Name, Total LEads Assigned, Converted Leads, (Sales + budget monthly)
        const usersResult = await findAllUsers({
          page: 1,
          limit: 100000,
        });

        const users = usersResult.data;

        const leads = await leadRepository.find({
          relations: {
            assigned_to: true,
            status: true,
          },
        });

        const convertedStatuses = new Set(["business done", "completed"]);

        const monthNames = [
          "Jan",
          "Feb",
          "Mar",
          "Apr",
          "May",
          "Jun",
          "Jul",
          "Aug",
          "Sep",
          "Oct",
          "Nov",
          "Dec",
        ];

        const currentMonthIndex = new Date().getMonth(); // Jan = 0, Jul = 6

        const staffPerformance = users.map((user: any) => {
          const report: any = {
            staffId: user.id,
            staffName:
              `${user.first_name ?? ""} ${user.last_name ?? ""}`.trim(),
          };

          // Only iterate till current month
          monthNames
            .slice(0, currentMonthIndex + 1)
            .forEach((month, monthIndex) => {
              const assignedLeads = leads.filter((lead: any) => {
                if (!lead.assigned_to) return false;

                return (
                  lead.assigned_to.id === user.id &&
                  new Date(lead.created_at).getMonth() === monthIndex
                );
              });

              const convertedLeads = assignedLeads.filter((lead: any) =>
                convertedStatuses.has(lead.status?.name?.trim().toLowerCase()),
              );

              const sales = convertedLeads.reduce(
                (sum: number, lead: any) => sum + Number(lead.budget || 0),
                0,
              );

              report[month] = {
                leadsAssigned: assignedLeads.length,
                convertedLeads: convertedLeads.length,
                sales,
              };
            });

          return report;
        });

        // Project snapshot (status counts)
        const projectSnapshot = {
          inProgress:
            projectStatusCounts.find((s: any) => s.status === "In Progress")
              ?.count || 0,
          completed:
            projectStatusCounts.find((s: any) => s.status === "Completed")
              ?.count || 0,
          completedProject:
            projectStatusCounts.find((s: any) => s.status === "Completed") ||
            [],
          allProject: projectStatusCounts,
          open:
            projectStatusCounts.find((s: any) => s.status === "Open")?.count ||
            0,
        };

        const monthWiseProjectRenewalData: Record<string, any[]> = {};

        for (const project of allProjects) {
          const renewalDate = project.renewal_date;

          if (!renewalDate) continue;

          const dateObj = new Date(renewalDate);
          const month = monthNames[dateObj.getMonth()];

          if (!monthWiseProjectRenewalData[month]) {
            monthWiseProjectRenewalData[month] = [];
          }

          const category = project.project_type?.name || "Other";

          let categoryGroup = monthWiseProjectRenewalData[month].find(
            (cat) => cat.category === category,
          );

          if (!categoryGroup) {
            categoryGroup = {
              category,
              projects: [],
            };
            monthWiseProjectRenewalData[month].push(categoryGroup);
          }

          // ✅ Calculate milestone completion % (considering support milestone logic and open tickets)
          const supportMilestones =
            project.milestones?.filter(
              (m) => m.name.toLowerCase() === "support",
            ) || [];
          const nonSupportMilestones =
            project.milestones?.filter(
              (m) => m.name.toLowerCase() !== "support",
            ) || [];

          // Check if all non-support milestones are completed
          const allNonSupportMilestonesCompleted =
            nonSupportMilestones.length > 0 &&
            nonSupportMilestones.every(
              (m) => m.status?.toLowerCase() === "completed",
            );

          // Check if support milestone has any open tickets
          const supportMilestoneHasOpenTickets = supportMilestones.some(
            (m) =>
              m.tickets &&
              m.tickets.some(
                (ticket) => ticket.status?.toLowerCase() === "open",
              ),
          );

          let completionPercentage = 0;

          // If all non-support milestones are completed AND no open tickets in support milestone, project is 100% complete
          if (
            (allNonSupportMilestonesCompleted &&
              !supportMilestoneHasOpenTickets) ||
            (supportMilestones.some(
              (m) => m.status?.toLowerCase() === "open",
            ) &&
              nonSupportMilestones.length === 0 &&
              !supportMilestoneHasOpenTickets)
          ) {
            completionPercentage = 100;
          } else {
            // Otherwise calculate based on non-support milestones only
            const totalMilestones = nonSupportMilestones.length || 0;
            const completedMilestones =
              nonSupportMilestones.filter(
                (m) => m.status?.toLowerCase() === "completed",
              ).length || 0;

            completionPercentage =
              totalMilestones > 0
                ? Math.round((completedMilestones / totalMilestones) * 100)
                : 0;
          }

          categoryGroup.projects.push({
            name: project.name,
            company_name:
              project.client?.company_name || project.client?.name || null,
            date: dateObj.toLocaleDateString("en-GB", {
              day: "2-digit",
              month: "long",
              year: "numeric",
            }),
            status: completionPercentage,
          });
        }

        // Expenses data (unchanged)
        const expenses = {
          weekly: weeklyChart,
          monthly: monthlyChart,
          yearly: yearlyChart,
        };

        // Lead analytics (status) and lead type, all periods
        const leadAnalytics = {
          weekly: leadStatusWeekly,
          monthly: leadStatusMonthly,
          yearly: leadStatusYearly,
        };
        const leadType = {
          weekly: leadTypeWeekly,
          monthly: leadTypeMonthly,
          yearly: leadTypeYearly,
        };

        res.status(200).json({
          status: "success",
          data: {
            stats,
            projectSnapshot,
            leadAnalytics,
            leadType,
            projectRenewalData: monthWiseProjectRenewalData,
            expenses,
            staffPerformance,
          },
        });
        return;
      }

      // Non-admin: only return limited stats
      // 1. My Task (count of all open, in process tasks assigned to user)
      // 2. Today Follow up (from getLeadStats)
      // 3. Project (count of projects where user is assigned to a milestone or task)
      // 4. Performance Ratio (completed tasks / total task assigned)
      const [
        leadStats,
        leadStatusDaily,
        allProjects,
        allTasksInSystem,
        todayFollowupsCount,
      ] = await Promise.all([
        leadService.getLeadStats(userId, role),
        leadService.getDailyLeadStats(userId, role),

        // Get all projects where user is assigned to a milestone or task
        projectService.getAllProject(userId, role),
        // Get all tasks in the system for total count
        (async () => {
          const { data } = await projectTaskService.getAllTasks(userId, role);
          return data;
        })(),
        clientFollowupService.getTodayFollowupsCount(userId, role),
      ]);

      const taskData = await projectTaskService.getUserTaskCounts(userId);

      // My Task: count of open and in process tasks (use taskData for consistency)
      const myTaskCount = taskData.pending + taskData.inProgress;
      // Performance Ratio: completed / total assigned
      const performanceRatio =
        taskData.total > 0
          ? `${Math.round((taskData.completed / taskData.total) * 100)}%`
          : "0%";
      // Project: count of projects where user is assigned
      const projectCount = allProjects.length;
      // Today Follow up
      const todayFollowups = todayFollowupsCount || 0;

      // Calculate total task counts from all tasks in system
      // Filter out any tasks that might be marked as deleted AND tasks without proper milestone relationships
      // Also filter out tasks whose milestones don't belong to existing projects
      const validMilestoneIds = new Set();
      allProjects.forEach((project: any) => {
        project.milestones?.forEach((milestone: any) => {
          validMilestoneIds.add(milestone.id);
        });
      });

      const activeTasksInSystem = allTasksInSystem.filter(
        (t: any) =>
          !t.deleted &&
          t.milestone &&
          t.milestone.id &&
          validMilestoneIds.has(t.milestone.id),
      );
      const totalTasksInSystem = activeTasksInSystem.length;

      // Use more flexible status matching to handle different case variations
      const completedTasksInSystem = activeTasksInSystem.filter(
        (t: any) => t.status && t.status.toLowerCase().includes("completed"),
      ).length;

      const openTasksInSystem = activeTasksInSystem.filter(
        (t: any) =>
          t.status &&
          (t.status.toLowerCase().includes("open") ||
            t.status.toLowerCase().includes("pending")),
      ).length;

      const inProgressTasksInSystem = activeTasksInSystem.filter(
        (t: any) => t.status && t.status.toLowerCase().includes("progress"),
      ).length;

      const approvalTasksInSystem = activeTasksInSystem.filter(
        (t: any) => t.status && t.status.toLowerCase().includes("approval"),
      ).length;

      const taskStat = {
        totalTasks: totalTasksInSystem,
        completedTasks: completedTasksInSystem,
        openTasks: openTasksInSystem,
        inprogressTasks: inProgressTasksInSystem,
        approvalTasks: approvalTasksInSystem,
      };

      // Only return the counts for the four stats
      res.status(200).json({
        status: "success",
        data: {
          myTaskCount,
          todayFollowups,
          projectCount,
          performanceRatio,
          taskStat,
          leadStatusDaily,
        },
      });
      return;
    } catch (error) {
      next(error);
    }
  };

  return {
    getDashboardSummary,
  };
};
