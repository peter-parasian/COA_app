namespace WpfApp1.Shared.Helpers
{
    public static class MathHelper
    {
        public static (System.Double Tensile, System.Double Elongation) CalculateTensileAndElongation(
            System.Double t1, System.Double t2, System.Double e1, System.Double e2)
        {
            System.Double finalTensile = 0;
            System.Double finalElongation = 0;

            System.Boolean hasT1 = t1 > 0;
            System.Boolean hasT2 = t2 > 0;
            System.Boolean hasE1 = e1 > 0;
            System.Boolean hasE2 = e2 > 0;

            if (hasT1 && hasT2 && hasE1 && hasE2)
            {
                if (t1 <= t2)
                {
                    finalTensile = t1;
                    finalElongation = e1;
                }
                else
                {
                    finalTensile = t2;
                    finalElongation = e2;
                }
            }
            else if (hasT1 && hasT2 && (hasE1 || hasE2))
            {
                finalTensile = System.Math.Min(t1, t2);
                finalElongation = System.Math.Max(e1, e2);
            }
            else if ((hasT1 || hasT2) && hasE1 && hasE2)
            {
                finalTensile = System.Math.Max(t1, t2);
                finalElongation = System.Math.Max(e1, e2);
            }
            else
            {
                finalTensile = System.Math.Max(t1, t2);
                finalElongation = System.Math.Max(e1, e2);
            }
            return (System.Math.Round(finalTensile, 2), System.Math.Round(finalElongation, 2));
        }
    }
}