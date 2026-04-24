import { redirect } from "react-router";

export const loader = async () => {
  throw redirect("/portal");
};

export default function IndexRoute() {
  return null;
}
