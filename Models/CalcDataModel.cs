using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace GB_NewCadPlus_IV.Models
{
    public class CalcDataModel
    {
        // 脱硫系统输入
        public double InletSO2 { get; set; }
        public double OutletSO2 { get; set; }
        public double AtmPressure { get; set; }
        public double OutletTemp { get; set; }
        public double StandardFlow { get; set; }
        public double StackVelocity { get; set; }
        public int EquipCount { get; set; }
        public double CaSRatio { get; set; }

        // 循环泵输入
        public double LiquidGasRatio { get; set; }
        public double ResidenceTime { get; set; }
        public int SprayLayers { get; set; }
        public double PumpHead { get; set; }
        public double MotorFactor { get; set; }
        public double SelectedPumpFlow { get; set; }
        public double MotorPower { get; set; }
        public int PumpSpareCount { get; set; }

        // 管道输入
        public double InletVelocity { get; set; }
        public double InletDiameter { get; set; }
        public double OutletVelocity { get; set; }
        public double OutletDiameter { get; set; }

        // 氧化风机输入
        public double FanHead { get; set; }
        public double FanMotorFactor { get; set; }
        public double SelectedFanFlow { get; set; }
        public double FanMotorPower { get; set; }
        public int FanRunningCount { get; set; }
        public int FanSpareCount { get; set; }
        public double FanOutletVelocity { get; set; }
        public double FanOutletDiameter { get; set; }

        // 管道介质/条件（字符串）
        public string InletMedium { get; set; }
        public string InletCondition { get; set; }
        public string OutletMedium { get; set; }
        public string OutletCondition { get; set; }
        public string FanOutletMedium { get; set; }
        public string FanOutletCondition { get; set; }

        // 石灰石浆液泵（远）
        public double LimePurity { get; set; }
        public double SlurryDensity { get; set; }
        public int FarSystemCount { get; set; }
        public int FarRunHours { get; set; }
        public int FarPumpCount { get; set; }
        public double FarPumpHead { get; set; }
        public double FarMotorFactor { get; set; }
        public double FarPumpEfficiency { get; set; }
        public double FarSelectedFlow { get; set; }
        public double FarMotorPower { get; set; }

        // 石灰石浆液泵（近）
        public int NearSystemCount { get; set; }
        public int NearRunHours { get; set; }
        public int NearPumpCount { get; set; }
        public double NearPumpHead { get; set; }
        public double NearMotorFactor { get; set; }
        public double NearPumpEfficiency { get; set; }
        public double NearSelectedFlow { get; set; }
        public double NearMotorPower { get; set; }

        // 石膏排出泵（远）
        public double GypsumDensity { get; set; }
        public int FarGypsumSystemCount { get; set; }
        public int FarGypsumRunHours { get; set; }
        public int FarGypsumPumpCount { get; set; }
        public double FarGypsumHead { get; set; }
        public double FarGypsumMotorFactor { get; set; }
        public double FarGypsumEfficiency { get; set; }
        public double FarGypsumSelectedFlow { get; set; }
        public double FarGypsumMotorPower { get; set; }

        // 石膏排出泵（近）
        public int NearGypsumSystemCount { get; set; }
        public int NearGypsumRunHours { get; set; }
        public int NearGypsumPumpCount { get; set; }
        public double NearGypsumHead { get; set; }
        public double NearGypsumMotorFactor { get; set; }
        public double NearGypsumEfficiency { get; set; }
        public double NearGypsumSelectedFlow { get; set; }
        public double NearGypsumMotorPower { get; set; }

        // 缓冲泵
        public double BufferRunHours { get; set; }
        public int BufferSystemCount { get; set; }
        public int BufferPumpCount { get; set; }
        public double BufferHead { get; set; }
        public double BufferMotorFactor { get; set; }
        public double BufferEfficiency { get; set; }
        public double BufferSelectedFlow { get; set; }
        public double BufferMotorPower { get; set; }

        // 滤液水泵（远）
        public double FiltrateDensity { get; set; }
        public int FiltrateFarSystemCount { get; set; }
        public int FiltrateFarPumpCount { get; set; }
        public double FiltrateFarHead { get; set; }
        public double FiltrateFarMotorFactor { get; set; }
        public double FiltrateFarEfficiency { get; set; }
        public double FiltrateFarSelectedFlow { get; set; }
        public double FiltrateFarMotorPower { get; set; }

        // 滤液水泵（近）
        public int FiltrateNearSystemCount { get; set; }
        public int FiltrateNearPumpCount { get; set; }
        public double FiltrateNearHead { get; set; }
        public double FiltrateNearMotorFactor { get; set; }
        public double FiltrateNearEfficiency { get; set; }
        public double FiltrateNearSelectedFlow { get; set; }
        public double FiltrateNearMotorPower { get; set; }

        // 事故泵（远）
        public double AccidentDensity { get; set; }
        public int AccidentFarPumpCount { get; set; }
        public double AccidentFarHead { get; set; }
        public double AccidentFarMotorFactor { get; set; }
        public double AccidentFarEfficiency { get; set; }
        public double AccidentFarSelectedFlow { get; set; }
        public double AccidentFarMotorPower { get; set; }

        // 事故泵（近）
        public int AccidentNearPumpCount { get; set; }
        public double AccidentNearHead { get; set; }
        public double AccidentNearMotorFactor { get; set; }
        public double AccidentNearEfficiency { get; set; }
        public double AccidentNearSelectedFlow { get; set; }
        public double AccidentNearMotorPower { get; set; }

        // 地坑泵
        public int SumpPumpCount { get; set; }
        public double SumpHead { get; set; }
        public double SumpMotorFactor { get; set; }
        public double SumpEfficiency { get; set; }
        public double SumpSelectedFlow { get; set; }
        public double SumpMotorPower { get; set; }
    }
}
