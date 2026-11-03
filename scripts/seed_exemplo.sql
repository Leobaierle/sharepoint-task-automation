-- Exemplo de carga da lista fixa de clientes de uma rotina mensal (dados fictícios).
begin;

delete from sharepoint_monthly_task_clients
where routine_key = 'dctfweb_mensal';

insert into sharepoint_monthly_task_clients (
  routine_key, cliente_id, cliente_nome, area_id, processo_id, active, notes
)
values
  ('dctfweb_mensal', 101, 'Empresa Alfa Ltda', 3, 103, true, 'Carga inicial da rotina'),
  ('dctfweb_mensal', 102, 'Empresa Beta Ltda', 3, 103, true, 'Carga inicial da rotina'),
  ('dctfweb_mensal', 103, 'Empresa Gama S/A', 3, 103, false, 'Inativa: fica fora da geração');

commit;
