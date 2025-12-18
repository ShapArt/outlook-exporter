from core import sla


def test_status_text_variants_ru_en():
    assert sla.status_text_to_code("OK") == sla.STATUS_RESPONDED
    assert sla.status_text_to_code("ОК") == sla.STATUS_RESPONDED
    assert sla.status_text_to_code("Закрыть") == sla.STATUS_RESOLVED
    assert sla.status_text_to_code("closed") == sla.STATUS_RESOLVED
    assert sla.status_text_to_code("Нужно время") == sla.STATUS_WAITING_CUSTOMER
    assert sla.status_text_to_code("waiting customer") == sla.STATUS_WAITING_CUSTOMER
    assert sla.status_text_to_code("ожидаем клиента") == sla.STATUS_WAITING_CUSTOMER
    assert sla.status_text_to_code("просрочка") == sla.STATUS_OVERDUE
    assert sla.status_text_to_code("неинтересно/спам") == sla.STATUS_NOT_INTERESTING
