import { getSP } from "./PnPConfig";
import "@pnp/sp/profiles";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/batching";
import "@pnp/sp/profiles";

export const addData = async (
  values: any,
  requesterMail: any,
  requesterName: any
) => {
  try {
    const sp = getSP();

    return await sp.web.lists
      .getByTitle("Approval_Request")
      .items.add({
        Customer: values.customer,
        Subject: values.subject,
        Product: values.product,
        SupportType: values.supportType,
        Contact: values.contact,
        RequesterMail: requesterMail,
        RequesterName: requesterName,
      })
      .then((res: any) => console.log("data submitted"))
      .catch((err: any) => console.error(err));
  } catch (error) {
    console.error("Error adding answer:", error);
    throw error;
  }
};

export const getMail = async () => {
  try {
    const sp = getSP();

    const items = await sp.web.lists.getByTitle("CrmMail").items.getAll();
    return items;
  } catch (error) {
    console.error("Error getting mail_id:", error);
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
