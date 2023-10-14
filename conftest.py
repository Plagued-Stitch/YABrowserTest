
import pytest


# функция предоставляет кастомный функционал для имплементации doc string комментария
def get_test_case_docstring(item):

    full_name = ''

    if item._obj.__doc__:

        name = str(item._obj.__doc__.split('.')[0]).strip()
        full_name = ' '.join(name.split())

        if hasattr(item, 'callspec'):
            params = item.callspec.params

            res_keys = sorted([k for k in params])

            res = ['{0}_"{1}"'.format(k, params[k]) for k in res_keys]

            full_name += ' Parameters ' + str(', '.join(res))
            full_name = full_name.replace(':', '')

    return full_name


# функция имплементирует doc string комментарий в консоль
def pytest_itemcollected(item):

    if item._obj.__doc__:
        item._nodeid = get_test_case_docstring(item)


# функция имплементирует doc string комментарий в консоль с параметром --collect-only
def pytest_collection_finish(session):

    if session.config.option.collectonly is True:

        for item in session.items:
            
            if item._obj.__doc__:
                full_name = get_test_case_docstring(item)
                print(full_name)

        pytest.exit('Done!')
