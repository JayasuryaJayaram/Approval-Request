import * as React from "react";
import styles from "./ApprovalRequest.module.scss";
import type { IApprovalRequestProps } from "./IApprovalRequestProps";
import { Form, Input, Button, Modal } from "antd";
import { useState } from "react";
import { addData, getMail, getUserData } from "../service/spService";
import { Requestmail } from "../RequestMail/mailTrigger";

const ApprovalRequest = (props: IApprovalRequestProps) => {
  const [form] = Form.useForm();
  const [isModalVisible, setIsModalVisible] = useState<boolean>(false);

  const onFinish = async (values: any) => {
    try {
      setIsModalVisible(true);

      let data = await getMail();
      let userData = await getUserData();

      let senderMail = data[0].mail_id;
      let requesterName = userData.DisplayName;
      let requesterMail = userData.Email;

      await addData(values, requesterMail, requesterName);
      await Requestmail(values, senderMail, requesterName, requesterMail);
      // console.log("Approval Requset sent to", senderMail);

      form.resetFields();
      console.log("Data submitted to SharePoint list");
    } catch (error) {
      console.error("Error submitting form:", error);
    }
  };

  const handleModalOk = () => {
    setIsModalVisible(false);
  };

  return (
    <>
      <div className={styles.formContainer}>
        <div className={styles.heading}>Approval Request</div>
        <Form
          form={form}
          onFinish={onFinish}
          className={styles.form}
          layout="vertical"
        >
          {/* Ant Design Form Fields */}

          <Form.Item
            name="customer"
            label="Customer"
            rules={[{ required: true, message: "Please enter customer name" }]}
          >
            <Input placeholder="Enter customer name" />
          </Form.Item>
          <Form.Item
            name="subject"
            label="Subject"
            rules={[{ required: true, message: "Please enter case subject" }]}
          >
            <Input placeholder="Enter case subject" />
          </Form.Item>
          <Form.Item
            name="product"
            label="Product"
            rules={[{ required: true, message: "Please enter product name" }]}
          >
            <Input placeholder="Enter product name" />
          </Form.Item>
          <Form.Item
            name="supportType"
            label="Support Type"
            rules={[{ required: true, message: "Please enter support type" }]}
          >
            <Input placeholder="Enter support type" />
          </Form.Item>
          <Form.Item
            name="contact"
            label="Contact"
            rules={[{ required: true, message: "Please enter contact" }]}
          >
            <Input placeholder="Enter contact" />
          </Form.Item>
          <Form.Item>
            <div>
              {/*//className={styles.popRow} */}
              <Button
                type="primary"
                htmlType="submit"
                className={styles.submitBtn}
              >
                Submit
              </Button>
              <Modal
                title="Successfully submitted!.."
                visible={isModalVisible}
                onCancel={handleModalOk}
                footer={[
                  <Button key="submit" type="primary" onClick={handleModalOk}>
                    OK
                  </Button>,
                ]}
              >
                <p>Your request is submitted.</p>
              </Modal>
            </div>
          </Form.Item>
        </Form>
      </div>
    </>
  );
};

export default ApprovalRequest;
