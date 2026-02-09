# Phone Dashboard Improvements Checklist

## Completed

- [x] **Public Holiday Exclusion** - Auto-fetch ACT holidays from API, exclude from stats (v1.2.0)
- [x] **Staff Rankings** - Medal badges for top 3 performers, green highlighting (v1.3.0)
- [x] **Version Badge** - Display version number in top right corner (v1.3.0)

---

## Tier 1: High Impact

- [ ] **PDF/Excel Export**
  - Export current view as PDF snapshot
  - Download filtered data as CSV for Excel analysis
  - Use jsPDF + html2canvas

- [ ] **Comparison Mode (This Week vs Last Week)**
  - Side-by-side comparison of two date ranges
  - Delta indicators (+5%, -10 calls, etc.)
  - Trend arrows on all metrics

- [ ] **Peak Hour Staffing Recommendations**
  - Analyse call volume by 30-min slot
  - Highlight understaffed periods (high missed % or wait time)
  - Show "Recommended Staff" based on target service level

---

## Tier 2: Medium Impact

- [ ] **Queue Performance Trends**
  - Line chart showing each queue's missed % over time
  - Identify consistently underperforming queues

- [ ] **Caller Frequency Analysis**
  - "Frequent Callers" table (callers with 3+ calls in period)
  - Flag potential problem cases or high-need patients
  - Distinguish new vs returning callers

- [ ] **Custom Service Level Target per Queue**
  - Per-queue SLA configuration
  - Show each queue's performance against its target

- [ ] **Browser Notifications**
  - Optional alerts for SLA breaches
  - Use browser Notifications API

---

## Quick Wins

- [ ] **Print Stylesheet** - Optimise layout for printing
- [ ] **Persistent Filters** - Save filter state to localStorage
- [ ] **Dark Mode** - Toggle for dark theme
- [ ] **Metric Tooltips** - Explain what each metric means on hover
- [ ] **Data Age Indicator** - Show how fresh the loaded data is

---

## Tier 3: Longer-Term (Requires Integration)

- [ ] **Appointment Outcome Tracking**
  - Integration with practice management system
  - Track "Call → Booking Conversion Rate"
  - Requires BP API or manual data import

- [ ] **Cost Per Call Analysis**
  - Calculate labour cost per call (staff wages ÷ calls handled)
  - Show cost impact of missed calls
  - Requires wage data

- [ ] **GP Impact View**
  - Show how call volume affects GP appointment availability
  - Correlate reception performance with GP satisfaction
  - Requires GP scheduling data

- [ ] **Real-time Live View**
  - Live call counter with current queue depth
  - Real-time service level indicator
  - Requires WebSocket integration with phone system

---

## Four Pillar Alignment

| Pillar | Best Features |
|--------|---------------|
| **1. Team Excellence** | Staff Rankings ✓, Peak Staffing |
| **2. GP Service** | GP Impact View, Export (reduces admin) |
| **3. Patient Experience** | Caller Frequency, Queue Trends |
| **4. Sustainable Enterprise** | Cost Analysis, Comparison Mode |

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0.0 | 2025-01 | Initial dashboard |
| 1.1.0 | 2026-01 | Date range picker, week selector, heatmap tabs |
| 1.2.0 | 2026-02 | Public holiday exclusion |
| 1.3.0 | 2026-02 | Staff rankings, version badge |
| 1.3.1 | 2026-02 | Fix OneDrive picker (corrupted script tag) |
