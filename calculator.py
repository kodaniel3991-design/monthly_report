"""
월차보고 시스템 - 계산 엔진
모든 수식은 기존 Excel 파일에서 역공학으로 추출
단위: 천원(KRW), 명, 시간, 대(판매수량)
"""

from dataclasses import dataclass, field
from typing import Optional


# ── 데이터 컨테이너 ────────────────────────────────────────────────────────

@dataclass
class FactoryPL:
    """공장별 손익 (단위: 천원)"""
    # 식별
    name: str = ""          # 김해, 부산, 울산, 김해2

    # 기본
    qty: float = 0          # 판매수량 (대)
    prod_amount: float = 0  # 생산금액

    # 매출액
    sales_prod: float = 0   # 생산품 매출
    sales_out: float = 0    # 외주품 매출

    # 변동비
    inv_diff: float = 0     # 제품재고증감차
    material: float = 0     # 재료비
    mfg_welfare: float = 0
    mfg_power: float = 0
    mfg_trans: float = 0
    mfg_repair: float = 0
    mfg_supplies: float = 0
    mfg_fee: float = 0
    mfg_other: float = 0
    selling_trans: float = 0   # 판매운반비
    merch_purchase: float = 0  # 상품매입

    # 고정비 - 노무비
    labor_salary: float = 0    # 급료
    labor_wage: float = 0      # 임금
    labor_bonus: float = 0     # 상여금
    labor_retire: float = 0    # 퇴충전입액
    labor_outsrc: float = 0    # 외주용역비

    # 고정비 - 인건비
    staff_salary: float = 0
    staff_bonus: float = 0
    staff_retire: float = 0

    # 고정비 - 제조경비
    fix_depr: float = 0        # 감가상각비
    fix_lease: float = 0       # 지급임차료
    fix_outsrc: float = 0      # 외주가공비
    fix_other: float = 0       # 기타경비

    # 영업외 (전사 공통 - 총괄에만 입력)
    non_op_income: float = 0
    non_op_expense: float = 0
    interest_income: float = 0
    interest_expense: float = 0

    # ── 계산 속성 ────────────────────────────────────────────────────────

    @property
    def sales(self) -> float:
        """매출액 = 생산품 + 외주품"""
        return self.sales_prod + self.sales_out

    @property
    def mfg_expense(self) -> float:
        """변동 제조경비 합계"""
        return (self.mfg_welfare + self.mfg_power + self.mfg_trans +
                self.mfg_repair + self.mfg_supplies + self.mfg_fee + self.mfg_other)

    @property
    def variable_cost(self) -> float:
        """변동비 = 재고증감차 + 재료비 + 제조경비 + 판매운반비 + 상품매입"""
        return (self.inv_diff + self.material + self.mfg_expense +
                self.selling_trans + self.merch_purchase)

    @property
    def labor_cost(self) -> float:
        """노무비 합계"""
        return (self.labor_salary + self.labor_wage + self.labor_bonus +
                self.labor_retire + self.labor_outsrc)

    @property
    def staff_cost(self) -> float:
        """인건비 합계"""
        return self.staff_salary + self.staff_bonus + self.staff_retire

    @property
    def fix_mfg_expense(self) -> float:
        """고정 제조경비"""
        return self.fix_depr + self.fix_lease + self.fix_outsrc + self.fix_other

    @property
    def fixed_cost(self) -> float:
        """고정비 = 노무비 + 인건비 + 제조경비"""
        return self.labor_cost + self.staff_cost + self.fix_mfg_expense

    @property
    def contribution_margin(self) -> float:
        """한계이익 = 매출액 - 변동비"""
        return self.sales - self.variable_cost

    @property
    def operating_profit(self) -> float:
        """영업이익 = 한계이익 - 고정비"""
        return self.contribution_margin - self.fixed_cost

    @property
    def ordinary_profit(self) -> float:
        """경상이익 = 영업이익 + 영업외수익 - 영업외비용"""
        return self.operating_profit + self.non_op_income - self.non_op_expense

    def pct(self, value: float) -> float:
        """매출액 대비 비율 (%)"""
        return round(value / self.sales * 100, 2) if self.sales else 0


@dataclass
class GroupPL:
    """그룹별 집계 (RKM = 김해+부산, HKMC = 울산+김해2)"""
    gimhae: FactoryPL = field(default_factory=FactoryPL)
    busan: FactoryPL = field(default_factory=FactoryPL)
    ulsan: FactoryPL = field(default_factory=FactoryPL)
    gimhae2: FactoryPL = field(default_factory=FactoryPL)

    def rkm(self) -> FactoryPL:
        """RKM = 김해 + 부산"""
        return _sum_factories("RKM", self.gimhae, self.busan)

    def hkmc(self) -> FactoryPL:
        """HKMC = 울산 + 김해2"""
        return _sum_factories("HKMC", self.ulsan, self.gimhae2)

    def total(self) -> FactoryPL:
        """전체 합계"""
        return _sum_factories("계", self.gimhae, self.busan, self.ulsan, self.gimhae2)


def _sum_factories(name: str, *factories: FactoryPL) -> FactoryPL:
    """여러 공장 합산"""
    result = FactoryPL(name=name)
    fields_to_sum = [
        "qty", "prod_amount",
        "sales_prod", "sales_out", "inv_diff",
        "material", "mfg_welfare", "mfg_power", "mfg_trans", "mfg_repair",
        "mfg_supplies", "mfg_fee", "mfg_other",
        "selling_trans", "merch_purchase",
        "labor_salary", "labor_wage", "labor_bonus", "labor_retire", "labor_outsrc",
        "staff_salary", "staff_bonus", "staff_retire",
        "fix_depr", "fix_lease", "fix_outsrc", "fix_other",
        "non_op_income", "non_op_expense", "interest_income", "interest_expense",
    ]
    for f in factories:
        for field_name in fields_to_sum:
            setattr(result, field_name,
                    getattr(result, field_name) + getattr(f, field_name))
    return result


# ── 노동생산성 계산 ────────────────────────────────────────────────────────

@dataclass
class LaborInput:
    """노동생산성 입력값"""
    # 인원 (명)
    mgmt_rkm: float = 0
    mgmt_hkmc: float = 0
    prod_rkm: float = 0
    prod_hkmc: float = 0
    # 근무시간
    work_hours_rkm: float = 0
    work_hours_hkmc: float = 0
    # 상여금 (천원)
    bonus_prod_rkm: float = 0
    bonus_prod_hkmc: float = 0
    # 퇴직급여 (천원)
    retire_mgmt_rkm: float = 0
    retire_mgmt_hkmc: float = 0
    retire_prod_rkm: float = 0
    retire_prod_hkmc: float = 0

    @property
    def total_employees(self) -> float:
        return self.mgmt_rkm + self.mgmt_hkmc + self.prod_rkm + self.prod_hkmc

    @property
    def prod_employees(self) -> float:
        return self.prod_rkm + self.prod_hkmc

    @property
    def rkm_employees(self) -> float:
        return self.mgmt_rkm + self.prod_rkm

    @property
    def hkmc_employees(self) -> float:
        return self.mgmt_hkmc + self.prod_hkmc

    @property
    def total_work_hours(self) -> float:
        return self.work_hours_rkm + self.work_hours_hkmc

    @property
    def rkm_ratio(self) -> float:
        """근무시간 기준 RKM 비율"""
        return self.work_hours_rkm / self.total_work_hours if self.total_work_hours else 0

    @property
    def hkmc_ratio(self) -> float:
        return 1 - self.rkm_ratio

    @property
    def retire_prod_total(self) -> float:
        return self.retire_prod_rkm + self.retire_prod_hkmc

    @property
    def retire_total(self) -> float:
        return (self.retire_mgmt_rkm + self.retire_mgmt_hkmc +
                self.retire_prod_rkm + self.retire_prod_hkmc)


@dataclass
class LaborProductivity:
    """노동생산성 지표 계산 결과"""
    label: str = ""
    # 입력
    sales: float = 0          # 매출액
    value_added: float = 0    # 부가가치액
    labor_cost: float = 0     # 노무비(급여+상여)
    retire_cost: float = 0    # 퇴직·복리비
    prod_amount: float = 0    # 생산금액
    employees: float = 0      # 상시종업원수
    prod_employees: float = 0 # 생산직 인원
    work_hours: float = 0     # 실작업시간
    retire_prod: float = 0    # 퇴직금(생산직)

    @property
    def value_added_ratio(self) -> float:
        """부가가치율 = 부가가치 / 매출액"""
        return self.value_added / self.sales if self.sales else 0

    @property
    def labor_productivity(self) -> float:
        """노동생산성 = 부가가치 / 종업원수 (천원/인)"""
        return self.value_added / self.employees if self.employees else 0

    @property
    def labor_income_ratio(self) -> float:
        """근로소득배분율 = 노무비 / 부가가치"""
        return self.labor_cost / self.value_added if self.value_added else 0

    @property
    def retire_ratio(self) -> float:
        """퇴직복리 배분율 = 퇴직비 / 부가가치"""
        return self.retire_cost / self.value_added if self.value_added else 0

    @property
    def total_personnel_ratio(self) -> float:
        """인건비 배분율 = (노무비+퇴직복리) / 부가가치"""
        return (self.labor_cost + self.retire_cost) / self.value_added if self.value_added else 0

    @property
    def labor_cost_to_sales(self) -> float:
        """매출대비 노무비율"""
        return self.labor_cost / self.sales if self.sales else 0

    @property
    def wage_per_person(self) -> float:
        """1인당 임금수준 = 노무비 / 종업원 (천원/인/월)"""
        return self.labor_cost / self.employees if self.employees else 0

    @property
    def retire_per_person(self) -> float:
        """1인당 퇴직금 (생산직 기준)"""
        return self.retire_prod / self.prod_employees if self.prod_employees else 0

    @property
    def hourly_wage(self) -> float:
        """시간당 임금 = 생산직 노무비 / 실작업시간 (천원/시간)"""
        # 생산직 노무비 = 전체 노무비 중 생산직 비율 적용
        # 단순화: 총 노무비로 계산 (파일에서 확인된 방식)
        return self.labor_cost / self.work_hours if self.work_hours else 0

    @property
    def prod_per_person(self) -> float:
        """1인당 월 생산금액 = 생산금액 / 생산직 인원 (천원/인)"""
        return self.prod_amount / self.prod_employees if self.prod_employees else 0

    @property
    def prod_per_won(self) -> float:
        """1원당 생산금액 = 생산금액 / 총노무비"""
        total_labor = self.labor_cost + self.retire_prod
        return self.prod_amount / total_labor if total_labor else 0


# ── 핵심 계산 함수 ─────────────────────────────────────────────────────────

def calc_value_added(pl: FactoryPL) -> float:
    """
    附加價値 = 賣出額 - 変動費 + 変動 複利厚生費
    = 한계이익 + 복리후생비(변동)

    복리후생비는 내부 인건비 성격이므로 외부구입비용에서 제외.
    기존 Excel 파일에서 역공학으로 검증 완료:
      전체: 4,848,177 - 3,703,139 + 11,210 = 1,156,248 ✓
      RKM:  3,503,150 - 2,608,461 + 9,905  = 904,594   ✓
      HKMC: 1,345,027 - 1,094,678 + 1,305  = 251,654   ✓
    """
    return pl.contribution_margin + pl.mfg_welfare


def calc_labor_productivity_total(
    pl_total: FactoryPL,
    labor: LaborInput,
    labor_cost_total: float,  # 손익실적의 노무비 합계
    retire_total: float,
) -> LaborProductivity:
    """총괄 노동생산성 계산"""
    va = calc_value_added(pl_total)
    return LaborProductivity(
        label="계",
        sales=pl_total.sales,
        value_added=va,
        labor_cost=labor_cost_total,
        retire_cost=retire_total,
        prod_amount=pl_total.prod_amount,
        employees=labor.total_employees,
        prod_employees=labor.prod_employees,
        work_hours=labor.total_work_hours,
        retire_prod=labor.retire_prod_total,
    )


def calc_labor_productivity_by_division(
    pl_rkm: FactoryPL,
    pl_hkmc: FactoryPL,
    labor: LaborInput,
    labor_cost_total: float,
) -> tuple[LaborProductivity, LaborProductivity]:
    """
    사업부별 노동생산성 계산
    노무비 RKM/HKMC 배분: 근무시간 비율로 1차 배분 후
    상여금은 실제 지급액으로 조정
    """
    # 근무시간 비율로 기본급 배분
    base_cost = labor_cost_total - (labor.bonus_prod_rkm + labor.bonus_prod_hkmc)
    base_rkm = base_cost * labor.rkm_ratio
    base_hkmc = base_cost * labor.hkmc_ratio

    labor_rkm = base_rkm + labor.bonus_prod_rkm
    labor_hkmc = base_hkmc + labor.bonus_prod_hkmc

    retire_rkm = labor.retire_mgmt_rkm + labor.retire_prod_rkm
    retire_hkmc = labor.retire_mgmt_hkmc + labor.retire_prod_hkmc

    va_rkm = calc_value_added(pl_rkm)
    va_hkmc = calc_value_added(pl_hkmc)

    lp_rkm = LaborProductivity(
        label="RKM",
        sales=pl_rkm.sales,
        value_added=va_rkm,
        labor_cost=labor_rkm,
        retire_cost=retire_rkm,
        prod_amount=pl_rkm.prod_amount,
        employees=labor.rkm_employees,
        prod_employees=labor.prod_rkm,
        work_hours=labor.work_hours_rkm,
        retire_prod=labor.retire_prod_rkm,
    )
    lp_hkmc = LaborProductivity(
        label="HKMC",
        sales=pl_hkmc.sales,
        value_added=va_hkmc,
        labor_cost=labor_hkmc,
        retire_cost=retire_hkmc,
        prod_amount=pl_hkmc.prod_amount,
        employees=labor.hkmc_employees,
        prod_employees=labor.prod_hkmc,
        work_hours=labor.work_hours_hkmc,
        retire_prod=labor.retire_prod_hkmc,
    )
    return lp_rkm, lp_hkmc


def build_factory_pl_from_db(data: dict, factory: str) -> FactoryPL:
    """DB 레코드에서 공장별 FactoryPL 생성"""
    suffix = f"_{factory}"
    return FactoryPL(
        name=factory,
        qty=data.get(f"qty{suffix}", 0) or 0,
        prod_amount=data.get(f"prod{suffix}", 0) or 0,
        sales_prod=data.get(f"sales_prod{suffix}", 0) or 0,
        sales_out=data.get(f"sales_out{suffix}", 0) or 0,
        inv_diff=data.get(f"inv_diff{suffix}", 0) or 0,
        material=data.get(f"material{suffix}", 0) or 0,
        mfg_welfare=data.get(f"mfg_welfare{suffix}", 0) or 0,
        mfg_power=data.get(f"mfg_power{suffix}", 0) or 0,
        mfg_trans=data.get(f"mfg_trans{suffix}", 0) or 0,
        mfg_repair=data.get(f"mfg_repair{suffix}", 0) or 0,
        mfg_supplies=data.get(f"mfg_supplies{suffix}", 0) or 0,
        mfg_fee=data.get(f"mfg_fee{suffix}", 0) or 0,
        mfg_other=data.get(f"mfg_other{suffix}", 0) or 0,
        selling_trans=data.get(f"selling_trans{suffix}", 0) or 0,
        merch_purchase=data.get(f"merch_purchase{suffix}", 0) or 0,
        labor_salary=data.get(f"labor_salary{suffix}", 0) or 0,
        labor_wage=data.get(f"labor_wage{suffix}", 0) or 0,
        labor_bonus=data.get(f"labor_bonus{suffix}", 0) or 0,
        labor_retire=data.get(f"labor_retire{suffix}", 0) or 0,
        labor_outsrc=data.get(f"labor_outsrc{suffix}", 0) or 0,
        staff_salary=data.get(f"staff_salary{suffix}", 0) or 0,
        staff_bonus=data.get(f"staff_bonus{suffix}", 0) or 0,
        staff_retire=data.get(f"staff_retire{suffix}", 0) or 0,
        fix_depr=data.get(f"fix_depr{suffix}", 0) or 0,
        fix_lease=data.get(f"fix_lease{suffix}", 0) or 0,
        fix_outsrc=data.get(f"fix_outsrc{suffix}", 0) or 0,
        fix_other=data.get(f"fix_other{suffix}", 0) or 0,
        non_op_income=data.get("non_op_income", 0) or 0,
        non_op_expense=data.get("non_op_expense", 0) or 0,
        interest_income=data.get("interest_income", 0) or 0,
        interest_expense=data.get("interest_expense", 0) or 0,
    )


def build_labor_input_from_db(data: dict) -> LaborInput:
    """DB 레코드에서 LaborInput 생성"""
    return LaborInput(
        mgmt_rkm=data.get("mgmt_rkm", 0) or 0,
        mgmt_hkmc=data.get("mgmt_hkmc", 0) or 0,
        prod_rkm=data.get("prod_rkm", 0) or 0,
        prod_hkmc=data.get("prod_hkmc", 0) or 0,
        work_hours_rkm=data.get("work_hours_rkm", 0) or 0,
        work_hours_hkmc=data.get("work_hours_hkmc", 0) or 0,
        bonus_prod_rkm=data.get("bonus_prod_rkm", 0) or 0,
        bonus_prod_hkmc=data.get("bonus_prod_hkmc", 0) or 0,
        retire_mgmt_rkm=data.get("retire_mgmt_rkm", 0) or 0,
        retire_mgmt_hkmc=data.get("retire_mgmt_hkmc", 0) or 0,
        retire_prod_rkm=data.get("retire_prod_rkm", 0) or 0,
        retire_prod_hkmc=data.get("retire_prod_hkmc", 0) or 0,
    )


if __name__ == "__main__":
    """
    2026년 3월 실제 RKM 합계값으로 검증
    (손익실적 파일 col11 기준)
    """
    rkm = FactoryPL(
        name="RKM",
        qty=12701,          prod_amount=1518529,
        sales_prod=1612418, sales_out=1890732,
        inv_diff=93889,     material=878368,
        mfg_welfare=9905,   mfg_power=31464,
        mfg_trans=1051,     mfg_repair=3288,
        mfg_supplies=5054,  mfg_fee=18664,
        mfg_other=11285,    selling_trans=10682,
        merch_purchase=1544811,
    )
    hkmc = FactoryPL(
        name="HKMC",
        qty=27804,          prod_amount=1179559,
        sales_prod=1163271, sales_out=181756,
        inv_diff=-16289,    material=879356,
        mfg_welfare=1305,   mfg_power=9006,
        selling_trans=7034, merch_purchase=213674,
    )
    total = _sum_factories("계", rkm, hkmc)

    print("=== 계산 검증 (2026년 3월 실적) ===")
    print(f"RKM  매출: {rkm.sales:>12,.0f}  기대: 3,503,150  {'✓' if abs(rkm.sales-3503150)<1 else '✗'}")
    print(f"HKMC 매출: {hkmc.sales:>12,.0f}  기대: 1,345,027  {'✓' if abs(hkmc.sales-1345027)<1 else '✗'}")
    print(f"전체 매출: {total.sales:>12,.0f}  기대: 4,848,177  {'✓' if abs(total.sales-4848177)<1 else '✗'}")
    print()
    va_rkm  = calc_value_added(rkm)
    va_hkmc = calc_value_added(hkmc)
    va_tot  = calc_value_added(total)
    print(f"RKM  부가가치: {va_rkm:>12,.0f}  기대:   904,594  {'✓' if abs(va_rkm-904594)<1 else '✗'}")
    print(f"HKMC 부가가치: {va_hkmc:>12,.0f}  기대:   251,654  {'✓' if abs(va_hkmc-251654)<5 else '✗'}")
    print(f"전체 부가가치: {va_tot:>12,.0f}  기대: 1,156,248  {'✓' if abs(va_tot-1156248)<5 else '✗'}")
    print()
    print(f"RKM  부가가치율: {va_rkm/rkm.sales*100:.2f}%  기대: 25.82%")
    print(f"HKMC 부가가치율: {va_hkmc/hkmc.sales*100:.2f}%  기대: 18.71%")
    print(f"전체 부가가치율: {va_tot/total.sales*100:.2f}%  기대: 23.85%")
