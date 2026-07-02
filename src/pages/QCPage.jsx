export function QCPage({ orders, onAdvance, logs, addLog }) {
  const qcOrders = orders.filter((order) => order.stage === 'qc');
  const getOrderLotCode = (order = {}) => order.lot || order.lotCode || order.orderCode || order.id || '-'

  const handlePass = (orderId) => {
    onAdvance(orderId, 'completed');
    addLog(`QC PASS lệnh ${orderId}, chuyển sang hoàn thành.`);
  };

  const handleReject = (orderId) => {
    addLog(`QC FAIL lệnh ${orderId}, yêu cầu kiểm tra lại.`);
  };

  return (
    <div className="page-content">
      <section className="panel">
        <h2>QC chất lượng</h2>
        <p className="panel-text">Kiểm tra chất lượng và cập nhật kết quả trước khi ghi nhật ký hoàn thành.</p>
        <div className="table-wrapper">
          <table>
            <thead>
              <tr>
                <th>Mã lô</th>
                <th>Sản phẩm</th>
                <th>Khối lượng</th>
                <th>Hành động</th>
              </tr>
            </thead>
            <tbody>
              {qcOrders.map((order) => (
                <tr key={order.id}>
                  <td>{getOrderLotCode(order)}</td>
                  <td>{order.product}</td>
                  <td>{order.quantityKg} kg</td>
                  <td className="action-row">
                    <button onClick={() => handlePass(order.id)} className="primary-button">
                      PASS
                    </button>
                    <button onClick={() => handleReject(order.id)} className="secondary-button">
                      FAIL
                    </button>
                  </td>
                </tr>
              ))}
              {qcOrders.length === 0 && (
                <tr><td colSpan="4" className="empty-row">Không có lệnh QC.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </section>
      <section className="panel small-panel">
        <h3>Nhật ký QC gần nhất</h3>
        <ul className="log-list">
          {logs.slice(-3).reverse().map((log) => (
            <li key={log.id}><strong>{log.time}</strong> {log.entry}</li>
          ))}
        </ul>
      </section>
    </div>
  );
}
