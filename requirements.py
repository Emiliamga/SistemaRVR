import pkg_resources
with open('requirements.txt', 'w') as f:
    for dist in pkg_resources.working_set:
        if not dist.key.startswith('pip'):  # exclui o próprio pip
            f.write(f'{dist.key}=={dist.version}\n')
