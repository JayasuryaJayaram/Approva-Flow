import { getSP } from "./PnPConfig";
import "@pnp/sp/profiles";

export const getData = async () => {
  try {
    const sp = getSP();

    const items = await sp.web.lists
      .getByTitle("Approval_Request")
      .items.getAll();
    return items;
  } catch (error) {
    console.error("Error getting mail_id:", error);
    throw error;
  }
};

export const updateApprovalStatus = async (itemId: any, status: string) => {
  try {
    const sp = getSP();

    return await sp.web.lists
      .getByTitle("Approval_Request")
      .items.getById(itemId)
      .update({
        ApprovalStatus: status,
      });
  } catch (error) {
    console.error("Error updating approval status:", error);
    throw error;
  }
};

export const getUserData = async () => {
  try {
    const sp = getSP();

    const details = await sp.profiles.myProperties();
    return details;
  } catch (error) {
    console.error("Error getting user info:", error);
    throw error;
  }
};
