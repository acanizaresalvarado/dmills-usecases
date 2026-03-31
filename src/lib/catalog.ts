import rawData from "../data/usecases.generated.json";

export type CoreCase = (typeof rawData)["coreCases"][number];
export type AddOn = (typeof rawData)["addOns"][number];
export type GroupInfo = (typeof rawData)["groups"][number];
export type SiteData = typeof rawData;

export const siteData = rawData as SiteData;

export const priorityOrder = ["Quick Win", "Strategic Bet", "Fill-In", "Trap"] as const;
export const recommendationOrder = ["First Wave", "Strategic Portfolio", "Keep Visible"] as const;
export const filterOrder = ["STRIP", "SCALE", "NORTH", "OCEAN"] as const;

export const prioritySlugMap: Record<string, string> = {
  "Quick Win": "quick-win",
  "Strategic Bet": "strategic-bet",
  "Fill-In": "fill-in",
  Trap: "trap",
};

export const slugPriorityMap = Object.fromEntries(
  Object.entries(prioritySlugMap).map(([label, slug]) => [slug, label]),
) as Record<string, string>;

function rank<T extends readonly string[]>(value: string, ordered: T) {
  const index = ordered.indexOf(value as T[number]);
  return index === -1 ? ordered.length : index;
}

export function sortCases(cases: CoreCase[]) {
  return [...cases].sort((left, right) => {
    return (
      rank(left.recommendation, recommendationOrder) - rank(right.recommendation, recommendationOrder) ||
      rank(left.priority, priorityOrder) - rank(right.priority, priorityOrder) ||
      right.valueScore - left.valueScore ||
      left.complexityScore - right.complexityScore ||
      left.code.localeCompare(right.code)
    );
  });
}

export function sortAddOns(addOns: AddOn[]) {
  return [...addOns].sort((left, right) => {
    return left.departmentLabel.localeCompare(right.departmentLabel) || left.title.localeCompare(right.title);
  });
}

export function getGroupBySlug(slug: string) {
  return siteData.groups.find((group) => group.slug === slug);
}

export function getCasesByGroup(slug: string) {
  return sortCases(siteData.coreCases.filter((item) => item.departmentSlug === slug));
}

export function getCasesByPriority(priority: string) {
  return sortCases(siteData.coreCases.filter((item) => item.priority === priority));
}

export function getCaseBySlug(slug: string) {
  return siteData.coreCases.find((item) => item.slug === slug);
}

export function getAddOnBySlug(slug: string) {
  return siteData.addOns.find((item) => item.slug === slug);
}

export function countByDepartment(cases: CoreCase[]) {
  return siteData.groups.map((group) => ({
    label: group.label,
    value: cases.filter((item) => item.departmentSlug === group.slug).length,
  }));
}

export function countByPriority(cases: CoreCase[]) {
  return priorityOrder.map((priority) => ({
    label: priority,
    value: cases.filter((item) => item.priority === priority).length,
  }));
}

export function formatPriorityLink(priority: string) {
  return `/priorities/${prioritySlugMap[priority] ?? priority.toLowerCase()}/`;
}
