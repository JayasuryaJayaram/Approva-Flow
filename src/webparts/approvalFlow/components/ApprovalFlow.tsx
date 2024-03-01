import * as React from "react";
import { useEffect, useState } from "react";
import type { IApprovalFlowProps } from "./IApprovalFlowProps";
import styles from "./ApprovalFlow.module.scss";
import { Table, Button, Modal } from "antd";
import {
  getData,
  updateApprovalStatus,
  getUserData,
} from "../service/spService";
import { ColumnsType } from "antd/es/table";
import { ApproveMail } from "../Mail/ApprovalMail";
import { RejectMail } from "../Mail/RejectMail";

interface DataType {
  key: string;
  customer: string;
  subject: string;
  product: string;
  supportType: string;
  contact: string;
  ApprovalStatus: string | null;
  requesterName: string;
  requesterMail: string;
}

const ApprovalFlow = (props: IApprovalFlowProps) => {
  const [data, setData] = useState<DataType[]>([]);
  const [isModalVisible, setIsModalVisible] = useState<boolean>(false);
  const [modalText, setModalText] = useState<string>("");

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    try {
      const requestDetails: any[] = await getData();
      const formattedData = requestDetails.map((item) => ({
        key: item.Id.toString(),
        customer: item.Customer,
        subject: item.Subject,
        product: item.Product,
        supportType: item.SupportType,
        contact: item.Contact,
        ApprovalStatus: item.ApprovalStatus,
        requesterName: item.RequesterName,
        requesterMail: item.RequesterMail,
      }));
      setData(formattedData);
      console.log(requestDetails);
    } catch (error) {
      console.error("Error fetching data:", error);
    }
  };

  const handleModalOk = () => {
    setIsModalVisible(false);
  };

  const showModal = (text: string) => {
    setModalText(text);
    setIsModalVisible(true);
  };

  const handleApprove = async (record: DataType) => {
    try {
      // Update local state first
      const updatedData = data.map((item) =>
        item.key === record.key ? { ...item, ApprovalStatus: "Approved" } : item
      );
      setData(updatedData);

      // Update SharePoint list item with new status
      await updateApprovalStatus(record.key, "Approved");

      let userData = await getUserData();
      let userName = userData.DisplayName;

      // Send approval mail
      await ApproveMail(record.requesterMail, record.requesterName, userName);

      showModal("Approved");
      // You might want to update the UI here after successful approval
    } catch (error) {
      console.error("Error approving:", error);
    }
  };

  const handleReject = async (record: DataType) => {
    try {
      // Update local state first
      const updatedData = data.map((item) =>
        item.key === record.key ? { ...item, ApprovalStatus: "Rejected" } : item
      );
      setData(updatedData);

      // Update SharePoint list item with new status
      await updateApprovalStatus(record.key, "Rejected");

      let userData = await getUserData();
      let userName = userData.DisplayName;

      // Send rejection mail
      await RejectMail(record.requesterMail, record.requesterName, userName);

      showModal("Rejected");
      // You might want to update the UI here after successful rejection
    } catch (error) {
      console.error("Error rejecting:", error);
    }
  };

  const columns: ColumnsType<DataType> = [
    {
      title: "Id",
      dataIndex: "key",
      key: "key",
    },
    {
      title: "Customer",
      dataIndex: "customer",
      key: "customer",
    },
    {
      title: "Subject",
      dataIndex: "subject",
      key: "subject",
    },
    {
      title: "Product",
      dataIndex: "product",
      key: "product",
    },
    {
      title: "Support Type",
      dataIndex: "supportType",
      key: "supportType",
    },
    {
      title: "Contact",
      dataIndex: "contact",
      key: "contact",
    },
    {
      title: "Approval",
      key: "actions",
      render: (text, record) => (
        <div className={styles.actions}>
          {record.ApprovalStatus === null && (
            <>
              <Button
                className={styles.approveButton}
                shape="round"
                onClick={() => handleApprove(record)}
              >
                Approve
              </Button>
              <Button
                className={styles.rejectButton}
                shape="round"
                onClick={() => handleReject(record)}
              >
                Reject
              </Button>
            </>
          )}
          {record.ApprovalStatus !== null && (
            <span>{record.ApprovalStatus}</span>
          )}
        </div>
      ),
    },
  ];

  var customStyles = `
       :where(.css-1rqnfsa).ant-table-wrapper .ant-table-container table>thead>tr:first-child >*:last-child  {
         text-align: center;
       }

       :where(.css-1rqnfsa).ant-btn.ant-btn-round.ant-btn {
        width: 25%;
       }
  `;

  return (
    <div className={styles.container}>
      <style>{customStyles}</style>
      <div className={styles.heading}>Approval Dashboard</div>
      <Table columns={columns} dataSource={data} />
      <Modal
        title={modalText}
        visible={isModalVisible}
        onCancel={handleModalOk}
        footer={[
          <Button key="submit" type="primary" onClick={handleModalOk}>
            OK
          </Button>,
        ]}
      >
        <p>Request is {modalText.toLowerCase()}.</p>
      </Modal>
    </div>
  );
};

export default ApprovalFlow;
