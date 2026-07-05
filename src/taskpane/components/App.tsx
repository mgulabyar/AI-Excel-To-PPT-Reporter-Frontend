import React, { useState } from "react";
import {
  Box,
  Typography,
  Button,
  Paper,
  CircularProgress,
  Fade,
  Stack,
  TextField,
  Card,
  Tabs,
  Tab,
} from "@mui/material";
import {
  QueryStats,
  TrendingDown,
  TrendingUp,
  Article,
  Lightbulb,
  PlayForWork,
  Assessment,
  AutoGraph,
} from "@mui/icons-material";
import axios from "axios";

interface KPI {
  label: string;
  value: string;
  status: string;
}
interface ReportData {
  report_title: string;
  kpis: KPI[];
  executive_summary: string;
  strategic_recommendations: string[];
}

const App: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const [tabValue, setTabValue] = useState(0);
  const [dataSummary, setDataSummary] = useState("");
  const [report, setReport] = useState<ReportData | null>(null);

  const generateReport = async () => {
    if (!dataSummary) return;
    setLoading(true);
    setReport(null);
    try {
      const res = await axios.post("http://127.0.0.1:8009/api/report/generate", {
        data_summary: dataSummary,
      });
      setReport(res.data.report);
      setTabValue(0);
    } catch (e) {
      alert("Error connecting to Engine 8009");
    }
    setLoading(false);
  };

  const syncToSlide = async () => {
    if (!report) return;
    try {
      await PowerPoint.run(async (context) => {
        const kpiText = report.kpis.map((k) => `${k.label}: ${k.value}`).join(" | ");
        const recsText = report.strategic_recommendations.join("\n• ");
        const finalContent = `${report.report_title.toUpperCase()}\n\n[KPIs]: ${kpiText}\n\n[SUMMARY]: ${report.executive_summary}\n\n[STRATEGY]:\n• ${recsText}`;

        Office.context.document.setSelectedDataAsync(
          finalContent,
          { coercionType: Office.CoercionType.Text },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Failed)
              alert("Select a text box first!");
          }
        );
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  };

  return (
    <Box
      sx={{
        height: "100vh",
        bgcolor: "#0a0a0a",
        display: "flex",
        flexDirection: "column",
        color: "#f8fafc",
      }}
    >
      <Paper
        elevation={0}
        sx={{
          p: 2,
          bgcolor: "#000000",
          borderBottom: "1px solid #1e293b",
          display: "flex",
          alignItems: "center",
          gap: 1.5,
        }}
      >
        <QueryStats sx={{ fontSize: 30, color: "#10b981" }} />
        <Typography
          variant="subtitle2"
          fontWeight={900}
          sx={{ fontSize: "16px", letterSpacing: 1, color:'#10b981' }}
        >
         AI Excel PPT Reporter
        </Typography>
      </Paper>

      <Box sx={{ p: 2, flexGrow: 1, overflowY: "auto" }}>
        <Typography
          variant="caption"
          fontWeight={900}
          color="#475569"
          sx={{ mb: 1, display: "block" }}
        >
          DATA SUMMARY INPUT
        </Typography>
        <TextField
          fullWidth
          multiline
          rows={3}
          placeholder="Paste Excel summary data here..."
          value={dataSummary}
          onChange={(e) => setDataSummary(e.target.value)}
          sx={{
            mb: 2,
            "& .MuiOutlinedInput-root": {
              fontSize: "13px",
              bgcolor: "#111827",
              color: "#fff",
              "& fieldset": { borderColor: "#1e293b" },
              "&:hover fieldset": { borderColor: "#10b981" },
            },
          }}
        />

        <Button
          fullWidth
          variant="contained"
          disableElevation
          onClick={generateReport}
          disabled={loading}
          sx={{
            bgcolor: "#111827",
            color: "#10b981",
            py: 1.5,
            fontWeight: 800,
            borderRadius: 2,
            border: "1px solid #1e293b",
            "&:hover": { bgcolor: "#1e293b" },
            "&.Mui-disabled": { bgcolor: "#1e293b", color: "#475569" },
          }}
        >
          {loading ? (
            <CircularProgress size={22} sx={{ color: "#10b981" }} />
          ) : (
            "GENERATE ANALYTICS"
          )}
        </Button>

        {report && (
          <Fade in={true}>
            <Box sx={{ mt: 3 }}>
              <Tabs
                value={tabValue}
                onChange={(_, v) => setTabValue(v)}
                variant="fullWidth"
                sx={{
                  mb: 2,
                  borderBottom: "1px solid #1e293b",
                  "& .MuiTab-root": { color: "#475569", fontSize: "11px", fontWeight: 800 },
                  "& .Mui-selected": { color: "#10b981 !important" },
                  "& .MuiTabs-indicator": { bgcolor: "#10b981" },
                }}
              >
                <Tab
                  label="INSIGHTS"
                  icon={<AutoGraph sx={{ fontSize: 18 }} />}
                  iconPosition="start"
                />
                <Tab
                  label="METRICS"
                  icon={<Assessment sx={{ fontSize: 18 }} />}
                  iconPosition="start"
                />
              </Tabs>

              {tabValue === 0 && (
                <Box>
                  <Card
                    variant="outlined"
                    sx={{
                      p: 2,
                      mb: 2,
                      bgcolor: "#111827",
                      borderColor: "#1e293b",
                      borderRadius: 3,
                    }}
                  >
                    <Stack direction="row" spacing={1} alignItems="center" sx={{ mb: 1 }}>
                      <Article sx={{ fontSize: 16, color: "#10b981" }} />
                      <Typography variant="caption" fontWeight={900} color="#10b981">
                        SUMMARY
                      </Typography>
                    </Stack>
                    <Typography
                      variant="body2"
                      sx={{ color: "#d1d5db", fontSize: "12px", lineHeight: 1.6 }}
                    >
                      {report.executive_summary}
                    </Typography>
                  </Card>

                  <Typography
                    variant="caption"
                    fontWeight={900}
                    color="#475569"
                    sx={{ mb: 1.5, display: "block" }}
                  >
                    STRATEGIC RECOMMENDATIONS
                  </Typography>
                  <Stack spacing={1} sx={{ mb: 3 }}>
                    {report.strategic_recommendations.map((rec, i) => (
                      <Paper
                        key={i}
                        variant="outlined"
                        sx={{ p: 1.5, bgcolor: "#000", borderColor: "#1e293b", borderRadius: 2 }}
                      >
                        <Stack direction="row" spacing={1.5}>
                          <Lightbulb sx={{ fontSize: 14, mt: 0.3, color: "#fbbf24" }} />
                          <Typography variant="body2" sx={{ color: "#f8fafc", fontSize: "11px" }}>
                            {rec}
                          </Typography>
                        </Stack>
                      </Paper>
                    ))}
                  </Stack>

                  <Button
                    fullWidth
                    variant="contained"
                    startIcon={<PlayForWork />}
                    onClick={syncToSlide}
                    sx={{
                      bgcolor: "#10b981",
                      color: "#000",
                      fontWeight: 800,
                      py: 1.2,
                      borderRadius: 2,
                      "&:hover": { bgcolor: "#059669" },
                    }}
                  >
                    SYNC TO SLIDE
                  </Button>
                </Box>
              )}

              {tabValue === 1 && (
                <Stack spacing={1.5}>
                  {report.kpis.map((kpi, i) => (
                    <Card
                      key={i}
                      variant="outlined"
                      sx={{
                        p: 2,
                        bgcolor: "#111827",
                        border: "1px solid #1e293b",
                        borderRadius: 3,
                      }}
                    >
                      <Stack direction="row" justifyContent="space-between" alignItems="center">
                        <Box>
                          <Typography variant="caption" color="#475569" fontWeight={800}>
                            {kpi.label.toUpperCase()}
                          </Typography>
                          <Typography
                            variant="h5"
                            fontWeight={900}
                            color={kpi.status === "Negative" ? "#f87171" : "#10b981"}
                          >
                            {kpi.value}
                          </Typography>
                        </Box>
                        {kpi.status === "Negative" ? (
                          <TrendingDown sx={{ fontSize: 24, color: "#f87171", opacity: 0.8 }} />
                        ) : (
                          <TrendingUp sx={{ fontSize: 24, color: "#10b981", opacity: 0.8 }} />
                        )}
                      </Stack>
                    </Card>
                  ))}
                </Stack>
              )}
            </Box>
          </Fade>
        )}
      </Box>

      {/* FO */}
      <Box sx={{ p: 1, textAlign: "center", borderTop: "1px solid #1e293b" }}>
        <Typography variant="caption" sx={{ color: "#334155", fontSize: "9px", fontWeight: 800 }}>
          ANALYSIS ENGINE NODE 8009
        </Typography>
      </Box>
    </Box>
  );
};

export default App;
