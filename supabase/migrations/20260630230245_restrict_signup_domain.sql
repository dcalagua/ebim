-- Solo cuentas del dominio de GRUPO EBIM pueden crear sesión (herramienta interna comercial)
create or replace function public.restrict_ebim_signup()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  if not (new.email ilike '%@ebim.pe' or new.email ilike '%@grupoebim.com') then
    raise exception 'Solo se permiten cuentas @ebim.pe o @grupoebim.com';
  end if;
  return new;
end;
$$;

drop trigger if exists restrict_ebim_signup_trigger on auth.users;
create trigger restrict_ebim_signup_trigger
  before insert on auth.users
  for each row execute function public.restrict_ebim_signup();
