/** A two-column table of rules: what, then how. */
export function RuleTable({ rows }: { rows: Array<[string, string]> }) {
  return (
    <table className="mini-table">
      <tbody>
        {rows.map(([k, v]) => (
          <tr key={k}>
            <th scope="row">{k}</th>
            <td>{v}</td>
          </tr>
        ))}
      </tbody>
    </table>
  )
}
